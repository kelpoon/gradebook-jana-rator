# Importing necessary modules and assets
from flask import Flask
from flask import request
from rubricReaderWriter import run_rubricReaderWriter # Importing my run_rubricReaderWriter function from rubricReaderWriter.py
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# Run gradebook route
@app.route("/run_gradebook")
def run_gradebook():
    folder_url = request.args.get('folder_url') # Gets the folder id for run_rubricReaderWriter to use
    gradebook_url = request.args.get('gradebook_url') # Gets the gradebook google doc id for run_rubricReaderWriter to use
    run_rubricReaderWriter(folder_url, gradebook_url) # Runs the run_rubricReaderWriter function
    return {"message": "Success!"} # if run_rubricReaderWriter runs it returns "Success" to the UI so the client knows the program ran

# Runs the app
if __name__ == "__main__":
    app.run(debug=True)