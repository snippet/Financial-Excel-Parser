import json
import os

import requests
from dotenv import load_dotenv
from flask import Flask, jsonify, request
from flask_cors import CORS

load_dotenv()

from openai import OpenAI
client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])

from processor import processor

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

UPLOAD_FOLDER = "uploads"
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({"error": "No selected file"}), 400

    if file:
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        processor(file.filename)

        response = {"filename": file.filename, "message": "File successfully uploaded"}
        return jsonify(response), 200


@app.route("/test", methods=["GET"])
def test_route():
    processor("example_0.xlsx")
    return jsonify({"message": "This is a test response"}), 200


@app.route("/files", methods=["GET"])
def list_files():
    files = os.listdir(UPLOAD_FOLDER)
    return jsonify({"files": files}), 200


@app.route("/chat", methods=["POST"])
def chat_with_gpt():
    data = request.get_json()
    query = data.get("query")
    files = data.get("files")

    if not query:
        return jsonify({"error": "Query is required"}), 400

    if not files:
        return jsonify({"error": "Files are required"}), 400

    financialData = []

    for file in files:
        file_name, _ = os.path.splitext(file)
        file_path = os.path.join("./parsed_files", f"{file_name}.json")
        if os.path.exists(file_path):
            with open(file_path, "r") as f:
                file_data = json.load(f)
                financialData.append(file_data)
        else:
            return jsonify({"error": f"File {file}.json not found"}), 404

    context = f"You are an expert financial analyst. Use you knowledge base to answer questions about audited financial statements. Knowledge Base: {json.dumps(financialData)}"
    print(context)

    api_key = os.getenv("OPENAI_API_KEY")

    if not api_key:
        return jsonify({"error": "OpenAI API key not found"}), 500

    chat_completion = client.chat.completions.create(
        messages=[
            {"role": "system", "content": context},
            {"role": "user", "content": query}
        ],
        model="gpt-4o",
    )

    response_text = chat_completion.choices[0].message.content

    return jsonify({"message": response_text}), 200

if __name__ == "__main__":
    app.run(debug=True)
