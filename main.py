import subprocess
import os
import traceback

def launch_electron():
    base_dir = os.path.dirname(os.path.abspath(__file__))

    electron_cmd = os.path.join(base_dir, 'node_modules', '.bin', 'electron.cmd')
    log_path = os.path.join(base_dir, 'debug_log.txt')

    with open(log_path, 'w', encoding='utf-8') as log:
        log.write("🟡 main.py started\n")
        log.write(f"🔹 Base dir: {base_dir}\n")
        log.write(f"🔹 electron.cmd path: {electron_cmd}\n")

        if not os.path.exists(electron_cmd):
            log.write("❌ electron.cmd not found. Run `npm install`.\n")
            return

        try:
            # Launch Electron from root using "."
            subprocess.Popen([electron_cmd, "."], cwd=base_dir)
            log.write("✅ Electron launched\n")
        except Exception:
            log.write("❌ Exception while launching Electron:\n")
            log.write(traceback.format_exc())

if __name__ == "__main__":
    launch_electron()