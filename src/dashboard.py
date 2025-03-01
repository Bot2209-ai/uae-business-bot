from flask import Flask, render_template_string, send_file
import json
import os

app = Flask(__name__)

def get_status():
    try:
        with open('status.json') as f:
            return json.load(f)
    except:
        return {'emails': 0, 'last_check': 'Never', 'errors': []}

@app.route('/')
def home():
    status = get_status()
    return render_template_string('''
        <h1>UAE Business Email Finder</h1>
        <p>Total Emails: {{ status.emails }}</p>
        <p>Last Check: {{ status.last_check }}</p>
        
        <h3>Recent Errors ({{ status.errors|length }}):</h3>
        <ul>
        {% for error in status.errors %}
            <li>{{ error }}</li>
        {% endfor %}
        </ul>
        
        <a href="/download"><button>Download Emails</button></a>
    ''', status=status)

@app.route('/download')
def download():
    return send_file('emails.csv', as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
