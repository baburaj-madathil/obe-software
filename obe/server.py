# Update imports at the top of obe/server.py
from .mapper import generate_co_wk_excel, generate_co_po_mapping

# ... [Keep the rest of your webserver code unchanged] ...

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
