# Update imports at the top of obe/server.py
from .mapper import generate_co_wk_excel, generate_co_po_mapping

import os
import shutil
import uuid
import json
import asyncio
import io
import sys
import socket
import threading
from pathlib import Path
from typing import AsyncGenerator

from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, StreamingResponse
from obe.mapper import generate_co_wk_excel, generate_co_po_mapping

# Configuration
MODEL_NAME = "gemini-2.5-flash"
PORT = 8756

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

# Ensure input and output directories exist in the current folder
INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


def get_local_ip() -> str:
    """Helper function to automatically detect the host machine's actual LAN IP address."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('8.8.8.8', 1))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return ip


class RealtimeStreamLogger(io.TextIOBase):
    """Intercepts stdout print statements and pushes them to an asyncio queue for streaming."""
    def __init__(self, queue: asyncio.Queue, loop: asyncio.AbstractEventLoop, original_stdout):
        self.queue = queue
        self.loop = loop
        self.original_stdout = original_stdout

    def write(self, s: str):
        self.original_stdout.write(s)
        self.original_stdout.flush()
        text = s.strip()
        if text:
            asyncio.run_coroutine_threadsafe(self.queue.put(text), self.loop)
        return len(s)

    def flush(self):
        self.original_stdout.flush()


class OBEService:
    """Service class responsible for executing OBE mappings with dynamic API keys."""
    def __init__(self, model_name: str):
        self.model_name = model_name

    def process_syllabus(self, file_path: Path, api_key: str, unique_co_po_filename: str, job_temp_dir: Path):
        """
        Executes CO-WK and CO-PO generation steps inside an isolated temp directory,
        using the user's provided API key.
        """
        original_cwd = os.getcwd()
        try:
            os.chdir(job_temp_dir)

            # Step 1: Generate intermediate CO-WK Mapping
            generate_co_wk_excel(
                pdf_path=str(file_path),
                model_name=self.model_name,
                api_key=api_key
            )

            # Step 2: Generate final CO-PO Mapping
            generate_co_po_mapping(
                model_name=self.model_name,
                api_key=api_key,
                input_excel="CO_WK_Mapping.xlsx"
            )

            # Locate generated file
            generated_file = None
            for fname in ["CO_PO_Mapping.xlsx", "Gemini_Improved_CO_PO_Mapping.xlsx"]:
                candidate = job_temp_dir / fname
                if candidate.exists():
                    generated_file = candidate
                    break

            if not generated_file:
                excels = list(job_temp_dir.glob("*.xlsx"))
                if excels:
                    generated_file = excels[-1]

            if generated_file and generated_file.exists():
                target_dest = OUTPUT_DIR / unique_co_po_filename
                shutil.copy(generated_file, target_dest)
            else:
                raise FileNotFoundError("CO-PO Mapping output file was not created by generator.")

        finally:
            os.chdir(original_cwd)


obe_service = OBEService(model_name=MODEL_NAME)
app = FastAPI(title="OBE Mapping Server")


@app.get("/", response_class=HTMLResponse)
async def main_page():
    return """
    <!DOCTYPE html>
    <html>
    <head>
        <title>OBE Mapping Generator</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; line-height: 1.6; background-color: #f4f6f9; }
            .container { max-width: 650px; margin: 0 auto; padding: 30px; background: #ffffff; border: 1px solid #ddd; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
            h2 { color: #333; margin-top: 0; }
            .form-group { margin-bottom: 20px; }
            label { font-weight: bold; display: block; margin-bottom: 5px; color: #555; }
            input[type="text"], input[type="file"] { width: 100%; padding: 10px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; }
            button { background-color: #007bff; color: white; border: none; padding: 12px 20px; font-size: 16px; border-radius: 4px; cursor: pointer; width: 100%; margin-top: 10px; }
            button:hover { background-color: #0056b3; }
            button:disabled { background-color: #cccccc; cursor: not-allowed; }
            
            #progress-container { display: none; margin-top: 25px; }
            .progress-box { width: 100%; background-color: #e0e0e0; border-radius: 6px; overflow: hidden; height: 20px; }
            .progress-bar { width: 0%; height: 100%; background-color: #007bff; transition: width 0.3s ease; }
            
            #log-box { display: none; margin-top: 15px; height: 180px; background: #1e1e1e; color: #4af626; font-family: monospace; font-size: 13px; padding: 12px; border-radius: 6px; overflow-y: auto; white-space: pre-wrap; }
            
            #result-container { display: none; margin-top: 25px; padding: 15px; background: #e9f5ff; border: 1px solid #b8daff; border-radius: 6px; }
            .download-btn { display: inline-block; margin-top: 10px; padding: 12px 20px; background-color: #28a745; color: white; text-decoration: none; border-radius: 4px; font-size: 15px; font-weight: bold; }
            .download-btn:hover { background-color: #218838; }
            .error-message { color: #dc3545; margin-top: 15px; word-break: break-all; }
        </style>
    </head>
    <body>
        <div class="container">
            <h2>Upload Syllabus PDF</h2>
            <p>Enter your Gemini API key and select your syllabus PDF file.</p>
            
            <form id="upload-form">
                <div class="form-group">
                    <label for="api-key-input">Gemini API Key:</label>
                    <input id="api-key-input" name="api_key" type="text" placeholder="Enter your Gemini API key" required>
                </div>

                <div class="form-group">
                    <label for="file-input">Syllabus PDF File:</label>
                    <input id="file-input" name="file" type="file" accept=".pdf" required>
                </div>

                <button type="submit" id="submit-btn">Process OBE Mapping</button>
            </form>

            <div id="progress-container">
                <div class="progress-box">
                    <div id="progress-bar" class="progress-bar"></div>
                </div>
            </div>

            <div id="log-box"></div>

            <div id="result-container">
                <h3 style="margin-top:0; color:#0c5460;">Processing Complete!</h3>
                <p>Your mapping report has been generated:</p>
                <a id="link-co-po" class="download-btn" href="#" download>Download CO-PO Mapping Excel</a>
            </div>

            <div id="error-box" class="error-message"></div>
        </div>

        <script>
            const form = document.getElementById('upload-form');
            const apiKeyInput = document.getElementById('api-key-input');
            const fileInput = document.getElementById('file-input');
            const submitBtn = document.getElementById('submit-btn');
            const progressContainer = document.getElementById('progress-container');
            const progressBar = document.getElementById('progress-bar');
            const logBox = document.getElementById('log-box');
            const resultContainer = document.getElementById('result-container');
            const errorBox = document.getElementById('error-box');

            form.addEventListener('submit', async (e) => {
                e.preventDefault();
                if (!fileInput.files.length || !apiKeyInput.value.trim()) return;

                errorBox.innerText = '';
                logBox.innerText = '';
                resultContainer.style.display = 'none';
                progressContainer.style.display = 'block';
                logBox.style.display = 'block';
                submitBtn.disabled = true;
                progressBar.style.width = '5%';

                const formData = new FormData();
                formData.append('api_key', apiKeyInput.value.trim());
                formData.append('file', fileInput.files[0]);

                try {
                    const response = await fetch('/process-stream/', {
                        method: 'POST',
                        body: formData
                    });

                    if (!response.ok) throw new Error('Upload failed.');

                    const reader = response.body.getReader();
                    const decoder = new TextDecoder();
                    let buffer = '';

                    while (true) {
                        const { done, value } = await reader.read();
                        if (done) break;

                        buffer += decoder.decode(value, { stream: true });
                        const lines = buffer.split('\\n\\n');
                        buffer = lines.pop();

                        for (const line of lines) {
                            if (line.startsWith('data: ')) {
                                const data = JSON.parse(line.replace('data: ', ''));

                                if (data.log) {
                                    logBox.innerText += data.log + '\\n';
                                    logBox.scrollTop = logBox.scrollHeight;

                                    if (data.log.includes('[Step 1/5]')) progressBar.style.width = '15%';
                                    else if (data.log.includes('[Step 2/5]')) progressBar.style.width = '25%';
                                    else if (data.log.includes('[Step 3/5]')) progressBar.style.width = '45%';
                                    else if (data.log.includes('[Step 4/5]')) progressBar.style.width = '60%';
                                    else if (data.log.includes('[Step 5/5]')) progressBar.style.width = '75%';
                                    else if (data.log.includes('Analyzing CO')) progressBar.style.width = '85%';
                                    else if (data.log.includes('Process completed successfully!')) progressBar.style.width = '100%';
                                }

                                if (data.status === 'SUCCESS') {
                                    progressBar.style.width = '100%';
                                    document.getElementById('link-co-po').href = data.download_link;
                                    
                                    setTimeout(() => {
                                        resultContainer.style.display = 'block';
                                        submitBtn.disabled = false;
                                    }, 400);
                                } else if (data.status === 'ERROR') {
                                    throw new Error(data.message);
                                }
                            }
                        }
                    }

                } catch (err) {
                    errorBox.innerText = 'Error: ' + err.message;
                    submitBtn.disabled = false;
                }
            });
        </script>
    </body>
    </html>
    """


@app.post("/process-stream/")
async def process_syllabus_stream(
    api_key: str = Form(...),
    file: UploadFile = File(...)
):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF files are allowed.")

    if not api_key.strip():
        raise HTTPException(status_code=400, detail="API Key is required.")

    # Unique identifier per request/user
    user_hash = uuid.uuid4().hex[:8]
    
    # Save PDF to input/ directory
    input_pdf_filename = f"Syllabus_{user_hash}.pdf"
    uploaded_pdf_path = INPUT_DIR / input_pdf_filename

    with open(uploaded_pdf_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Isolated processing workspace inside input/
    job_temp_dir = INPUT_DIR / f"temp_{user_hash}"
    job_temp_dir.mkdir(parents=True, exist_ok=True)

    # Output file destination inside output/
    output_excel_filename = f"CO_PO_Mapping_{user_hash}.xlsx"

    queue = asyncio.Queue()
    loop = asyncio.get_running_loop()

    async def event_generator() -> AsyncGenerator[str, None]:
        original_stdout = sys.stdout
        sys.stdout = RealtimeStreamLogger(queue, loop, original_stdout)

        task = loop.run_in_executor(
            None, 
            obe_service.process_syllabus, 
            uploaded_pdf_path, 
            api_key.strip(),
            output_excel_filename,
            job_temp_dir
        )

        try:
            while not task.done() or not queue.empty():
                try:
                    log_text = await asyncio.wait_for(queue.get(), timeout=0.2)
                    yield f"data: {json.dumps({'log': log_text})}\n\n"
                except asyncio.TimeoutError:
                    await asyncio.sleep(0.05)

            task.result()

            success_payload = {
                "status": "SUCCESS",
                "download_link": f"/download/{output_excel_filename}"
            }
            yield f"data: {json.dumps(success_payload)}\n\n"

        except Exception as e:
            error_payload = {"status": "ERROR", "message": str(e)}
            yield f"data: {json.dumps(error_payload)}\n\n"

        finally:
            sys.stdout = original_stdout
            # Clean up temporary processing folder
            if job_temp_dir.exists():
                shutil.rmtree(job_temp_dir, ignore_errors=True)

    return StreamingResponse(event_generator(), media_type="text/event-stream")


@app.get("/download/{filename}")
async def download_file(filename: str):
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail=f"Requested output file '{filename}' was not found.")

    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def listen_for_quit(server):
    """Monitors console input in a separate thread for 'q' key press to terminate server gracefully."""
    while True:
        try:
            key = input()
            if key.strip().lower() == 'q':
                print("\n'q' pressed. Shutting down OBE Server gracefully...")
                server.should_exit = True
                break
        except (KeyboardInterrupt, EOFError):
            server.should_exit = True
            break


if __name__ == "__main__":
    import uvicorn

    host_ip = get_local_ip()

    print("\n" + "=" * 65)
    print("           OBE Mapping Server Started Successfully           ")
    print("=" * 65)
    print(f"  Access URL: http://{host_ip}:{PORT}")
    print(f"  Input Folder:  {INPUT_DIR}")
    print(f"  Output Folder: {OUTPUT_DIR}")
    print("=" * 65)
    print("  Press 'q' then ENTER in this window at any time to stop.")
    print("=" * 65 + "\n")

    config = uvicorn.Config(app=app, host=host_ip, port=PORT, log_level="info")
    server = uvicorn.Server(config)

    quit_thread = threading.Thread(target=listen_for_quit, args=(server,), daemon=True)
    quit_thread.start()

    server.run()


# Add this entry function at the bottom of obe/server.py
def run_server():
    """Entry point for the obe-server CLI command."""
    import uvicorn
    host_ip = get_local_ip()

    print("\n" + "=" * 65)
    print("           OBE Mapping Server Started Successfully           ")
    print("=" * 65)
    print(f"  Access URL: http://{host_ip}:{PORT}")
    print(f"  Input Folder:  {INPUT_DIR}")
    print(f"  Output Folder: {OUTPUT_DIR}")
    print("=" * 65)
    print("  Press 'q' then ENTER in this window at any time to stop.")
    print("=" * 65 + "\n")

    config = uvicorn.Config(app=app, host=host_ip, port=PORT, log_level="info")
    server = uvicorn.Server(config)

    quit_thread = threading.Thread(target=listen_for_quit, args=(server,), daemon=True)
    quit_thread.start()

    server.run()


if __name__ == "__main__":
    run_server()
