import threading
import queue
import time
import os
import base64
import io
import re
from datetime import datetime
from difflib import get_close_matches
import sys

import dash
from dash import dcc, html, Input, Output, State, callback_context, no_update
import dash_bootstrap_components as dbc
import pandas as pd
import speech_recognition as sr
import pyttsx3
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# -------------------------------------------------------------------------
# GLOBAL STATE
# -------------------------------------------------------------------------
class AppState:
    def __init__(self):
        self.price_df = None
        self.items = []  # List of (item, qty, rate, total)
        self.logs = []   # List of strings for log window
        self.is_listening = False
        self.stop_event = threading.Event()
        self.speech_queue = queue.Queue()
        self.customer_name = "-"  # Voice-driven customer name

    def add_log(self, text):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.logs.append(f"[{timestamp}] {text}")
        if len(self.logs) > 20:
            self.logs.pop(0)

state = AppState()

# -------------------------------------------------------------------------
# TTS & UTILITIES
# -------------------------------------------------------------------------
engine = pyttsx3.init()
engine.setProperty('rate', 160)

def talk(text):
    def _speak():
        try:
            local_engine = pyttsx3.init()
            local_engine.say(text)
            local_engine.runAndWait()
        except:
            pass
    threading.Thread(target=_speak, daemon=True).start()

def parse_quantity_and_item(text):
    text = text.lower().strip()
    words_to_num = {'one':1,'two':2,'three':3,'four':4,'five':5,'six':6,'seven':7,'eight':8,'nine':9,'ten':10}
    m = re.match(r'^\s*(\d+)\s+(.+)$', text)
    if m: return int(m.group(1)), m.group(2).strip()
    m2 = re.match(r'^\s*([a-z]+)\s+(.+)$', text)
    if m2 and m2.group(1) in words_to_num:
        return words_to_num[m2.group(1)], m2.group(2).strip()
    return 1, text

def fuzzy_lookup(item_text, price_df, cutoff=0.6):
    if price_df is None: return None, None
    item_text = item_text.strip().lower()
    if item_text in price_df.index:
        return item_text, float(price_df.loc[item_text, 'price'])
    matches = get_close_matches(item_text, list(price_df.index), n=1, cutoff=cutoff)
    if matches:
        return matches[0], float(price_df.loc[matches[0], 'price'])
    return None, None

def generate_bill_pdf(filename_pdf, bill_info):
    c = canvas.Canvas(filename_pdf, pagesize=A4)
    w, h = A4
    left = 20 * mm
    top = h - 20 * mm
    y = top

    c.setFont("Helvetica-Bold", 14)
    c.drawString(left, y, bill_info.get('store_name', 'STORE'))
    y -= 8 * mm
    c.setFont("Helvetica", 10)
    c.drawString(left, y, f"Date: {bill_info['date_time'].strftime('%Y-%m-%d %H:%M:%S')}")
    y -= 5 * mm
    c.drawString(left, y, f"Customer: {bill_info.get('customer_name','-')}    Phone: {bill_info.get('customer_phone','-')}")
    y -= 8 * mm
    c.line(left, y, w - left, y)
    y -= 6 * mm
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left, y, "Item"); c.drawString(left+70*mm, y, "Qty"); c.drawString(left+95*mm, y, "Rate"); c.drawString(left+120*mm, y, "Total")
    y -= 6 * mm
    c.line(left, y, w - left, y)
    y -= 6 * mm
    c.setFont("Helvetica", 10)
    
    for it, qty, rate, line_total in bill_info['items']:
        c.drawString(left, y, str(it))
        c.drawString(left+70*mm, y, str(qty))
        c.drawString(left+95*mm, y, f"{rate:.2f}")
        c.drawString(left+120*mm, y, f"{line_total:.2f}")
        y -= 6 * mm

    y -= 4 * mm
    c.line(left, y, w - left, y)
    y -= 8 * mm
    
    subtotal = sum(x[3] for x in bill_info['items'])
    gst_amt = subtotal * bill_info['gst_percent'] / 100
    disc_amt = subtotal * bill_info['discount_percent'] / 100
    total = subtotal + gst_amt - disc_amt
    
    c.drawRightString(w - left, y, f"Subtotal: {subtotal:.2f}")
    y -= 6 * mm
    c.drawRightString(w - left, y, f"GST: {gst_amt:.2f}")
    y -= 6 * mm
    c.drawRightString(w - left, y, f"Discount: -{disc_amt:.2f}")
    y -= 8 * mm
    c.setFont("Helvetica-Bold", 12)
    c.drawRightString(w - left, y, f"TOTAL: {total:.2f}")
    y -= 12 * mm
    c.setFont("Helvetica", 9)
    c.drawString(left, y, "Thank you for shopping!")
    c.save()

# -------------------------------------------------------------------------
# BACKGROUND LISTENER THREAD
# -------------------------------------------------------------------------
def background_listener():
    r = sr.Recognizer()
    mic = sr.Microphone()
    with mic as source:
        r.adjust_for_ambient_noise(source, duration=1)
    state.add_log("Listener thread started. Waiting for speech...")
    while not state.stop_event.is_set():
        try:
            with mic as source:
                audio = r.listen(source, phrase_time_limit=5)
            text = r.recognize_google(audio).lower().strip()
            state.speech_queue.put(text)
        except sr.UnknownValueError:
            continue
        except Exception as e:
            state.add_log(f"Error: {e}")
            time.sleep(1)

# -------------------------------------------------------------------------
# DASH APP
# -------------------------------------------------------------------------
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.DARKLY])
app.title = "Voice Bill Generator"

app.layout = dbc.Container([
    dcc.Interval(id='interval-component', interval=1000, n_intervals=0),
    dcc.Store(id='download-path-store'),

    dbc.Row([
        dbc.Col(html.H2("Voice-Assisted Bill Generator"), width=9, className="my-4"),
        dbc.Col(html.Img(src="/assets/logo.jpg", style={"height":"50px", "float":"right"}), width=3, className="my-4 text-end")
    ]),

    dbc.Card([
        dbc.CardBody([
            dbc.Row([
                dbc.Col([
                    dcc.Upload(
                        id='upload-data',
                        children=dbc.Button("📂 Load Price Excel", color="primary", className="w-100"),
                        multiple=False
                    ),
                    html.Div(id='output-filename', className="text-muted small mt-1")
                ], width=3),
                dbc.Col([
                    dbc.Button("🎤 Start Listening", id='btn-start', color="success", className="me-2"),
                    dbc.Button("⛔ Stop", id='btn-stop', color="danger", disabled=True)
                ], width=6, className="text-center"),
                dbc.Col([
                    dbc.Button("🖨️ Print Bill", id='btn-print', color="dark", className="w-100")
                ], width=3),
            ])
        ])
    ], className="mb-3"),

    dbc.Card([dbc.CardHeader("Live Speech Log"),
              dbc.CardBody(html.Div(id='live-logs', style={"height": "150px", "overflowY": "scroll", "fontFamily": "monospace", "backgroundColor": "#343a40", "color":"#ffffff", "padding": "10px"}))],
             className="mb-3"),

    dbc.Row([
        dbc.Col([html.H5("Current Items"), html.Div(id='table-container'), html.Hr(), html.H4(id='bill-summary', className="text-end text-primary")])
    ])
], fluid=True)

# -------------------------------------------------------------------------
# CALLBACKS
# -------------------------------------------------------------------------
@app.callback(
    Output('output-filename', 'children'),
    Input('upload-data', 'contents'),
    State('upload-data', 'filename')
)
def load_excel(contents, filename):
    if contents is None: return ""
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    try:
        df = pd.read_excel(io.BytesIO(decoded))
        df.columns = [c.lower().strip() for c in df.columns]
        if 'item' in df.columns and 'price' in df.columns:
            df['item'] = df['item'].astype(str).str.lower().str.strip()
            df = df.set_index('item')
            state.price_df = df
            state.add_log(f"Loaded {filename} successfully.")
            talk("Price list loaded.")
            return f"Loaded: {filename}"
        else:
            return "Error: Missing 'item' or 'price' columns"
    except Exception as e:
        return f"Error loading file: {e}"

@app.callback(
    Output('btn-start', 'disabled'),
    Output('btn-stop', 'disabled'),
    Input('btn-start', 'n_clicks'),
    Input('btn-stop', 'n_clicks'),
    prevent_initial_call=True
)
def toggle_listening(start_clicks, stop_clicks):
    ctx = callback_context
    if not ctx.triggered: return no_update, no_update
    button_id = ctx.triggered[0]['prop_id'].split('.')[0]

    if button_id == 'btn-start':
        if not state.is_listening:
            state.stop_event.clear()
            threading.Thread(target=background_listener, daemon=True).start()
            state.is_listening = True
            talk("Listening started")
            return True, False
    elif button_id == 'btn-stop':
        if state.is_listening:
            state.stop_event.set()
            state.is_listening = False
            talk("Listening stopped")
            return False, True
    return no_update, no_update

@app.callback(
    Output('live-logs', 'children'),
    Output('table-container', 'children'),
    Output('bill-summary', 'children'),
    Input('interval-component', 'n_intervals'),
)
def poll_updates(n):
    while not state.speech_queue.empty():
        text = state.speech_queue.get().lower().strip()
        
        # Check if the user said the name
        if text.startswith("name "):
            name = text.replace("name ", "").strip()
            state.customer_name = name
            state.add_log(f"Customer name set to {name}")
            talk(f"Hello {name}")
            continue

        # Otherwise parse items
        qty, item_name = parse_quantity_and_item(text)
        if state.price_df is None:
            state.add_log("Error: Price list not loaded!")
            talk("Please load price list first")
        else:
            match, price = fuzzy_lookup(item_name, state.price_df)
            if match:
                total = qty * price
                state.items.append((match, qty, price, total))
                state.add_log(f"Added: {qty} x {match}")
                talk(f"Added {qty} {match}")
            else:
                state.add_log(f"Unknown item: {item_name}")
                talk(f"Could not find {item_name}")

    # Logs
    log_elements = [html.Div(l) for l in reversed(state.logs)]

    # Table
    table_header = [html.Thead(html.Tr([html.Th("Item"), html.Th("Qty"), html.Th("Rate"), html.Th("Total")]))]
    table_body = [html.Tr([html.Td(i[0]), html.Td(i[1]), html.Td(f"{i[2]:.2f}"), html.Td(f"{i[3]:.2f}")]) for i in state.items]
    table = dbc.Table(table_header + [html.Tbody(table_body)], striped=True, bordered=True, hover=True, dark=True)

    # Summary
    subtotal = sum(x[3] for x in state.items)
    gst_amt = 0
    disc_amt = 0
    if state.items:
        gst_amt = subtotal * 0 / 100
        disc_amt = subtotal * 0 / 100
    final_total = subtotal + gst_amt - disc_amt
    summary_text = f"Sub: {subtotal:.2f} | GST: {gst_amt:.2f} | Disc: {disc_amt:.2f} | Total: {final_total:.2f}"

    return log_elements, table, summary_text

@app.callback(
    Output('download-path-store', 'data'),
    Input('btn-print', 'n_clicks'),
    prevent_initial_call=True
)
def print_bill(n):
    if not state.items:
        state.add_log("Cannot print empty bill.")
        return no_update

    filename = f"Bill_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    bill_data = {
        'store_name': 'ASB Cafeteria',
        'customer_name': state.customer_name,
        'customer_phone': '-',
        'items': state.items,
        'gst_percent': 0,
        'discount_percent': 0,
        'date_time': datetime.now()
    }
    
    try:
        generate_bill_pdf(filename, bill_data)
        state.add_log(f"PDF Generated: {filename}")
        talk("Bill generated successfully")
        
        # Open file (Windows/Linux)
        if sys.platform == 'win32':
            os.startfile(filename)
        else:
            import subprocess
            subprocess.call(['xdg-open', filename])
        
        # Clear items and logs after printing
        state.items = []
        state.logs = []
        state.customer_name = "-"
        talk("Bill cleared. Ready for a new bill.")
            
    except Exception as e:
        state.add_log(f"Print error: {e}")

    return None

# -------------------------------------------------------------------------
if __name__ == '__main__':
    app.run(debug=True, threaded=True)
