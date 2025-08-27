from flask import Flask

app = Flask(__name__)

@app.route('/')
def test():
    return "Flask is working!"

@app.route('/upload', methods=['POST'])  
def test_upload():
    return {"message": "Upload route works"}

if __name__ == '__main__':
    print("Starting minimal Flask test...")
    app.run(debug=True, port=5001)
