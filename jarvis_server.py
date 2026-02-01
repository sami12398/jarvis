#!/usr/bin/env python3
"""
J.A.R.V.I.S. Flask API Server
REST API for voice and web interfaces
"""

from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from jarvis_core import jarvis
import os
import threading
import time

app = Flask(__name__, static_folder='.')
CORS(app)

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/api/command', methods=['POST'])
def command():
    data = request.json
    cmd = data.get('command', '')
    
    print(f"\n[Command] {cmd}")
    result = jarvis.process_command(cmd)
    print(f"[Result] {result['message']}")
    
    return jsonify(result)

@app.route('/api/status', methods=['GET'])
def status():
    return jsonify({
        "status": "online",
        "version": "2.0",
        "timestamp": time.time()
    })

@app.route('/api/history', methods=['GET'])
def history():
    # Return last 10 commands (would need storage in production)
    return jsonify([])

def run_server(port=5000):
    print(f"Starting JARVIS Server on http://localhost:{port}")
    print("Press Ctrl+C to stop\n")
    app.run(host='0.0.0.0', port=port, debug=False, use_reloader=False)

if __name__ == '__main__':
    run_server()