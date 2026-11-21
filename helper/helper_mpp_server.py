# -*- coding: utf-8 -*-
from __future__ import annotations
import json, re, sys, traceback, os
from pathlib import Path
from urllib.parse import urlparse, unquote
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HOST = '127.0.0.1'
PORT = int(os.environ.get('GLC_HELPER_PORT', '8765'))
BASE_DIR = Path(__file__).resolve().parent
EXPORT_DIR = BASE_DIR / 'exports'
EXPORT_DIR.mkdir(exist_ok=True)

try:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore
    HAS_COM = True
    IMPORT_ERROR = ''
except Exception as e:
    HAS_COM = False
    IMPORT_ERROR = str(e)

def json_response(handler, status, payload):
    data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
    handler.send_response(status)
    handler.send_header('Content-Type', 'application/json; charset=utf-8')
    handler.send_header('Content-Length', str(len(data)))
    handler.send_header('Access-Control-Allow-Origin', '*')
    handler.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    handler.send_header('Access-Control-Allow-Headers', 'Content-Type')
    handler.end_headers()
    handler.wfile.write(data)

def send_file(handler, path: Path):
    if not path.exists() or not path.is_file():
        return json_response(handler, 404, {'ok': False, 'error': 'Arquivo não encontrado'})
    data = path.read_bytes()
    handler.send_response(200)
    handler.send_header('Content-Type', 'application/octet-stream')
    handler.send_header('Content-Length', str(len(data)))
    handler.send_header('Content-Disposition', f'attachment; filename="{path.name}"')
    handler.send_header('Access-Control-Allow-Origin', '*')
    handler.end_headers()
    handler.wfile.write(data)

def safe_filename(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]+', '_', str(name).strip())
    name = re.sub(r'\s+', '_', name)
    return name[:180] or 'projeto'

def parse_duration_to_days(value: str) -> int:
    """Converte '5 dias', '5d', '0d' em inteiro. Aceita 0 (marco)."""
    text = str(value or '').strip().lower()
    m = re.search(r'(\d+)', text)
    if not m:
        return 1
    return max(0, int(m.group(1)))

def build_mpp(payload: dict) -> Path:
    if not HAS_COM:
        raise RuntimeError('pywin32/MS Project não disponível neste ambiente. Instale Microsoft Project e pywin32 no Windows. ' + IMPORT_ERROR)

    records = payload.get('records') or []
    if not records:
        raise ValueError('Nenhum registro recebido para exportação do MPP.')

    project_name = safe_filename(payload.get('project_name') or 'PROJETO')
    out_path = EXPORT_DIR / f'{project_name}.mpp'

    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.Dispatch('MSProject.Application')
        app.Visible = False
        app.FileNew()
        project = app.ActiveProject
        inicio = str(payload.get('inicio') or '').strip()
        if inicio:
            try:
                project.ProjectStart = inicio.replace('-', '/')
            except Exception:
                pass
        previous_level = 1
        previous_task = None
        task_by_source_id = {}

        for rec in records:
            name = str(rec.get('nome') or '').strip() or 'Tarefa'
            level = int(rec.get('nivel') or 1)
            duration_days = parse_duration_to_days(rec.get('duracao') or '')
            task = project.Tasks.Add(name)
            if previous_task is not None:
                diff = level - previous_level
                if diff > 0:
                    for _ in range(diff):
                        app.SelectRow(previous_task.ID + 1, False, 1)
                        app.OutlineIndent()
                elif diff < 0:
                    for _ in range(abs(diff)):
                        app.SelectRow(task.ID, False, 1)
                        app.OutlineOutdent()
            try:
                if rec.get('marco'):
                    task.Duration = 0
                    task.Milestone = True
                elif rec.get('duracao'):
                    task.Duration = f'{duration_days}d' if duration_days > 0 else 0
            except Exception:
                pass
            try:
                if rec.get('text1'):
                    task.Text1 = str(rec.get('text1'))
                if rec.get('notas'):
                    task.Notes = str(rec.get('notas'))
            except Exception:
                pass
            task_by_source_id[str(rec.get('id'))] = task
            previous_task = task
            previous_level = level

        for rec in records:
            pred = str(rec.get('predecessoras') or '').strip()
            if not pred:
                continue
            task = task_by_source_id.get(str(rec.get('id')))
            if not task:
                continue
            preds = []
            for part in re.split(r'[;,]', pred):
                part = part.strip()
                if not part:
                    continue
                src_task = task_by_source_id.get(part)
                if src_task:
                    preds.append(str(src_task.ID))
            if preds:
                try:
                    task.Predecessors = ';'.join(preds)
                except Exception:
                    pass

        if out_path.exists():
            out_path.unlink()
        app.FileSaveAs(str(out_path))
        if not out_path.exists() or out_path.stat().st_size == 0:
            raise RuntimeError('O Microsoft Project não gerou o arquivo .MPP no disco.')
        return out_path
    finally:
        try:
            if app is not None:
                app.FileCloseEx(0)
                app.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

class Handler(BaseHTTPRequestHandler):
    server_version = 'HelperMPP/2.1'

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == '/health':
            return json_response(self, 200, {
                'ok': True,
                'message': f'Helper ativo em http://{HOST}:{PORT}',
                'has_com': HAS_COM,
                'import_error': IMPORT_ERROR,
                'export_dir': str(EXPORT_DIR)
            })
        if parsed.path == '/diag':
            return json_response(self, 200, {
                'ok': True,
                'python': sys.version,
                'cwd': os.getcwd(),
                'has_com': HAS_COM,
                'import_error': IMPORT_ERROR,
                'export_dir': str(EXPORT_DIR),
                'port': PORT
            })
        if parsed.path.startswith('/download/'):
            fname = unquote(parsed.path.split('/download/', 1)[1])
            return send_file(self, EXPORT_DIR / fname)
        return json_response(self, 404, {'ok': False, 'error': 'Rota não encontrada'})

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path != '/export/mpp':
            return json_response(self, 404, {'ok': False, 'error': 'Rota não encontrada'})
        try:
            length = int(self.headers.get('Content-Length', '0'))
            body = self.rfile.read(length)
            payload = json.loads(body.decode('utf-8'))
            saved = build_mpp(payload)
            return json_response(self, 200, {
                'ok': True,
                'file_name': saved.name,
                'download_url': f'http://{HOST}:{PORT}/download/{saved.name}'
            })
        except Exception as exc:
            return json_response(self, 500, {
                'ok': False,
                'error': str(exc),
                'trace': traceback.format_exc(limit=8)
            })

    def log_message(self, format, *args):
        sys.stdout.write('%s - - [%s] %s\n' % (self.client_address[0], self.log_date_time_string(), format % args))

def main():
    print(f'Helper MPP iniciado em http://{HOST}:{PORT}')
    if not HAS_COM:
        print('AVISO: pywin32/MS Project não detectado. O endpoint /health responde, mas /export/mpp falhará até instalar o pywin32 e MS Project.')
        print('Detalhe:', IMPORT_ERROR)
    httpd = ThreadingHTTPServer((HOST, PORT), Handler)
    httpd.serve_forever()

if __name__ == '__main__':
    main()
