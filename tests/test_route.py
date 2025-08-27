#!/usr/bin/env python3
"""Simple test to check Flask routes"""

import requests
import json

# Test if the route exists by making a POST request
url = 'http://127.0.0.1:5000/upload'
data = {'employee_name': 'ostardo'}

try:
    response = requests.post(url, data=data)
    print(f"Status Code: {response.status_code}")
    print(f"Response: {response.text}")
except Exception as e:
    print(f"Error: {e}")
