import json
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
import tkinter as tk
from tkinter import filedialog, scrolledtext
import pythoncom
import xlwings as xw

PORT = 5050
state = {
    "app": None,
    "wb": None,
    "path": None,
    "log_fn": None,
    "root": None,
}


def log(msg: str):
    if state["log_fn"]:
        state["log_fn"](msg)


def run_on_main(fn):
    done = threading.Event()
    result = {}

    def wrapper():
        try:
            result["value"] = fn()
        except Exception as e:
            result["error"] = e
        finally:
            done.set()

    state["root"].after(0, wrapper)
    done.wait(timeout=30)
    if "error" in result:
        raise result["error"]
    return result.get("value")


def get_wb(book_name=None):
    if book_name and state["app"] is not None:
        return state["app"].books[book_name]
    return state["wb"]


class BridgeHandler(BaseHTTPRequestHandler):
    def _send_json(self, code: int, data: dict):
        body = json.dumps(data, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length) if length else b"{}"
        return json.loads(raw.decode("utf-8"))

    def log_message(self, format, *args):
        pass  # Suppress default stderr HTTP logs

    def do_GET(self):
        if self.path == "/status":
            if state["app"] is None:
                return self._send_json(200, {"connected": False, "books": []})
            try:
                info = run_on_main(lambda: {
                    "connected": len(state["app"].books) > 0,
                    "active": state["wb"].name if state["wb"] else None,
                    "books": [
                        {
                            "name": b.name,
                            "path": b.fullname,
                            "sheets": [s.name for s in b.sheets],
                        }
                        for b in state["app"].books
                    ],
                })
                return self._send_json(200, info)
            except Exception as e:
                return self._send_json(500, {"error": str(e)})
        self._send_json(404, {"error": "Not found"})

    def do_POST(self):
        if state["wb"] is None:
            return self._send_json(400, {"error": "Belum ada file Excel yang dibuka di ExcelAI."})

        try:
            data = self._read_json()
            book_name = data.get("book")

            if self.path == "/read":
                val = run_on_main(lambda: get_wb(book_name).sheets[data["sheet"]].range(data["range"]).value)
                log(f"READ {data['sheet']}!{data['range']}")
                return self._send_json(200, {"value": val})

            if self.path == "/write":
                def do_write():
                    wb = get_wb(book_name)
                    ws = wb.sheets[data["sheet"]]
                    ws.range(data["cell"]).value = data["value"]
                    if data.get("save", True):
                        wb.save(wb.fullname)
                run_on_main(do_write)
                log(f"WRITE {data['sheet']}!{data['cell']}")
                return self._send_json(200, {"ok": True})

            if self.path == "/add_sheet":
                def do_add():
                    wb = get_wb(book_name)
                    name = data["name"]
                    if name not in [s.name for s in wb.sheets]:
                        wb.sheets.add(name, after=wb.sheets[-1])
                    wb.save(wb.fullname)
                    return [s.name for s in wb.sheets]
                sheets = run_on_main(do_add)
                log(f"ADD SHEET: {data['name']}")
                return self._send_json(200, {"ok": True, "sheets": sheets})

            if self.path == "/run_code":
                def do_exec():
                    wb = get_wb(book_name)
                    local_vars = {"wb": wb, "xw": xw}
                    exec(data["code"], {}, local_vars)
                    if data.get("save", True):
                        wb.save(wb.fullname)
                run_on_main(do_exec)
                log("RUN CODE executed")
                return self._send_json(200, {"ok": True})

            self._send_json(404, {"error": "Unknown endpoint"})
        except Exception as e:
            log(f"ERROR: {e}")
            self._send_json(500, {"error": str(e)})


def start_server():
    server = HTTPServer(("127.0.0.1", PORT), BridgeHandler)
    server.serve_forever()


def main():
    root = tk.Tk()
    state["root"] = root
    root.title("ExcelAI Bridge")
    root.geometry("540x400")

    status_lbl = tk.Label(root, text=f"Server aktif di http://127.0.0.1:{PORT}", fg="green", font=("Segoe UI", 10, "bold"))
    status_lbl.pack(pady=(10, 4))

    file_lbl = tk.Label(root, text="Belum ada file Excel dibuka", fg="#333", font=("Segoe UI", 9))
    file_lbl.pack(pady=4)

    log_box = scrolledtext.ScrolledText(root, height=15, font=("Consolas", 9))
    log_box.pack(fill="both", expand=True, padx=12, pady=8)

    def append_log(msg: str):
        root.after(0, lambda: (log_box.insert(tk.END, msg + "\n"), log_box.see(tk.END)))

    state["log_fn"] = append_log

    def open_excel():
        paths = filedialog.askopenfilenames(
            filetypes=[
                ("Spreadsheet & Data", "*.xlsx *.xlsm *.xls *.xlsb *.csv *.tsv *.ods"),
                ("Excel Files", "*.xlsx *.xlsm *.xls *.xlsb"),
                ("CSV / TSV Files", "*.csv *.tsv"),
                ("All Files", "*.*"),
            ]
        )
        if not paths:
            return
        try:
            if state["app"] is not None:
                try:
                    if len(state["app"].books) == 0:
                        state["app"].quit()
                        state["app"] = None
                except Exception:
                    state["app"] = None
            if state["app"] is None:
                state["app"] = xw.App(visible=True)
                try:
                    for addin in state["app"].api.AddIns:
                        if addin.Installed:
                            state["app"].api.Workbooks.Open(addin.FullName)
                except Exception:
                    pass
            state["app"].visible = True
            for path in paths:
                wb = state["app"].books.open(path)
                state["wb"] = wb
                state["path"] = path
                append_log(f"Dibuka: {path}")
            if "Book1" in [b.name for b in state["app"].books] and len(state["app"].books) > 1:
                state["app"].books["Book1"].close()
            state["app"].activate(steal_focus=True)
            refresh_status()
        except Exception as e:
            append_log(f"Gagal membuka: {e}")

    def refresh_status():
        try:
            if state["app"] is not None:
                names = [b.name for b in state["app"].books]
                if names:
                    file_lbl.config(text="Terhubung: " + ", ".join(names), fg="#1F4E78")
                else:
                    state["wb"] = None
                    file_lbl.config(text="Belum ada file Excel dibuka", fg="#666")
            else:
                file_lbl.config(text="Belum ada file Excel dibuka", fg="#666")
        except Exception:
            state["app"] = None
            state["wb"] = None
            file_lbl.config(text="Belum ada file Excel dibuka (Excel ditutup)", fg="#b00020")
        root.after(2000, refresh_status)

    btn = tk.Button(root, text="Pilih & Buka File Excel", command=open_excel, bg="#1F4E78", fg="white", font=("Segoe UI", 10, "bold"), padx=12, pady=6)
    btn.pack(pady=(4, 8))

    threading.Thread(target=start_server, daemon=True).start()
    append_log(f"ExcelAI Bridge berjalan di http://127.0.0.1:{PORT}")
    refresh_status()
    root.mainloop()


if __name__ == "__main__":
    main()

