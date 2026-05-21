# NexConvert Tool

NexConvert is a Flask web app for converting documents between PDF and Word formats.

- PDF to DOCX conversion uses `pdf2docx`.
- DOCX/DOC to PDF conversion uses Microsoft Word automation on Windows and LibreOffice in Linux/Docker.
- Uploaded and converted files are stored temporarily and cleaned up automatically.

## Project Structure

```text
.
|-- app.py
|-- static/
|-- templates/
|-- uploads/
|-- converted/
|-- Dockerfile
`-- requirements.txt
```

## Run Locally

Create and activate a virtual environment:

```bash
python -m venv .venv
.venv\Scripts\activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Start the app:

```bash
python app.py
```

Open:

```text
http://127.0.0.1:5000
```

## Docker

Build the image:

```bash
docker build -t nexconvert-tool .
```

Run the container:

```bash
docker run -p 5000:5000 -e PORT=5000 nexconvert-tool
```

Open:

```text
http://127.0.0.1:5000
```

## Railway Deployment

This repository includes a `Dockerfile`, so Railway can deploy it as a Docker-based service.

1. Create a new Railway project.
2. Connect the GitHub repository: `pankajlp/nexconvert-tool`.
3. Railway will detect the Dockerfile and build the service.
4. No custom start command is required.

The app listens on Railway's `PORT` environment variable through Gunicorn.

## Notes

- DOCX/DOC to PDF conversion in Railway depends on LibreOffice, which is installed by the Dockerfile.
- Runtime uploads and converted files are intentionally ignored by Git.
- Files older than 30 minutes are cleaned from the upload and converted folders.
