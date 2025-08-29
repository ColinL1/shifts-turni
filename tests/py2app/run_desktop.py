#!/usr/bin/env python3
import sys
import os
import threading
import time
import webview
import socket
from app import app

def is_port_in_use(port):
    """Check if a port is already in use"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('localhost', port)) == 0

if __name__ == "__main__":
    # Set working directory appropriately for bundled app
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle
        bundle_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        base_dir = os.path.dirname(os.path.abspath(sys.executable))
        print(f"Running as frozen app from: {base_dir}")
        os.chdir(base_dir)
    else:
        # Running in normal Python environment
        bundle_dir = os.path.dirname(os.path.abspath(__file__))
        print(f"Running in development mode from: {bundle_dir}")

    # Define the port for the Flask app
    port = 5000
    
    # Check if the port is already in use
    if is_port_in_use(port):
        print(f"Port {port} is already in use. Please close the application using that port and try again.")
        sys.exit(1)

    # Define Flask server thread
    def start_flask():
        # Start the Flask app directly
        print(f"Starting Flask server on port {port}...")
        app.run(host='127.0.0.1', port=port, use_reloader=False)

    # Start Flask in a separate thread
    flask_thread = threading.Thread(target=start_flask)
    flask_thread.daemon = True
    flask_thread.start()
    
    # Wait for Flask to start
    print("Waiting for Flask server to start...")
    time.sleep(2)
    
    # Try to connect to verify the server is running
    max_attempts = 10
    attempt = 0
    server_ready = False
    
    while attempt < max_attempts:
        if is_port_in_use(port):
            server_ready = True
            break
        attempt += 1
        time.sleep(0.5)
    
    # Create and start webview window if server started successfully
    if server_ready:
        print("Flask server started successfully!")
        
        # Set window title based on whether we're running in a bundle or not
        if getattr(sys, 'frozen', False):
            window_title = "Shift Manager"
        else:
            window_title = "Shift Manager (Development)"
            
        # Create and start webview window
        webview.create_window(window_title, f"http://127.0.0.1:{port}/")
        webview.start()
        
        print("Window closed. Shutting down...")
    else:
        print("Failed to start Flask server. Please try again.")
