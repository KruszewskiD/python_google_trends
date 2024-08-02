from flask import Flask, request, send_file, render_template_string
from pytrends.request import TrendReq
import pandas as pd
import io
import time
import logging

app = Flask(__name__)

logging.basicConfig(level=logging.DEBUG, filename='app.log', filemode='w')

@app.route('/')
def index():
    return render_template_string('''
    <!DOCTYPE html>
    <html>
    <head>
        <title>PyTrends App</title>
        <script>
            function addInput() {
                var newInput = document.createElement("input");
                newInput.setAttribute("type", "text");
                newInput.setAttribute("name", "keywords");
                var form = document.getElementById("keywordsForm");
                form.appendChild(newInput);
                form.appendChild(document.createElement("br"));
            }
        </script>
    </head>
    <body>
        <h1>PyTrends App</h1>
        <form id="keywordsForm" action="/get_trends" method="post">
            <button type="button" onclick="addInput()">Add more keywords</button><br>
            <label for="language">Select language:</label>
            <select name="language" id="language">
                <option value="pl-PL">Polski</option>
                <option value="en-US">English</option>
                <option value="de-DE">Deutsch</option>
                <!-- Dodaj inne opcje językowe, jeśli są potrzebne -->
            </select><br>
            <input type="submit" value="Get Trends">
            <input type="text" name="keywords"><br>
        </form>
    </body>
    </html>
    ''')

def get_trends_data(keywords, language):
    attempts = 5
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"

    for attempt in range(attempts):
        try:
            logging.debug(f'Attempt {attempt+1} to fetch trends data with keywords: {keywords} and language: {language}')
            headers = {'User-Agent': user_agent}
            logging.debug(f'Using headers: {headers}')
            pytrends = TrendReq(hl=language, tz=360, requests_args={'headers': headers})
            pytrends.build_payload(keywords, cat=0, timeframe='today 12-m')
            data = pytrends.interest_over_time()
            if not data.empty:
                return data
            else:
                logging.debug('Data list is empty, retrying...')
                time.sleep(60)  # Opóźnienie 60 sekund przed kolejną próbą
        except Exception as e:
            logging.error(f'Request failed: {e}')
            if attempt < attempts - 1:
                logging.debug('Waiting for 60 seconds before the next attempt...')
                time.sleep(60)  # Opóźnienie 60 sekund przed kolejną próbą
            else:
                raise e

@app.route('/get_trends', methods=['POST'])
def get_trends():
    keywords = request.form.getlist('keywords')
    language = request.form.get('language', 'pl-PL')  # Domyślnie ustawiamy język na polski
    
    try:
        # Pobieranie danych trendów z obsługą ponawiania prób
        data = get_trends_data(keywords, language)
        
        # Remove 'isPartial' column if it exists
        if 'isPartial' in data.columns:
            data = data.drop(columns=['isPartial'])
        
        # Zachowanie oryginalnego formatu daty
        data.index = data.index.strftime('%Y-%m-%d %H:%M:%S')

        # Create a bytes buffer for the Excel file
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            data.to_excel(writer, sheet_name='Trends')

        buffer.seek(0)
        return send_file(buffer, as_attachment=True, download_name='trends.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return f"An error occurred: {e}", 500

if __name__ == '__main__':
    app.run(debug=True)
