"""
Сервер на Flask: общение по HTTP.
Нужно:
- alg.py лежит на сервере;
- ты в веб-админке жмёшь кнопку;
- клиенты, которые опрашивают сервер, видят новый запуск, СКАЧИВАЮТ alg.py и запускают его локально.

Запуск сервера: .venv/bin/python server.py
"""

import os
from io import BytesIO
from flask import Flask, jsonify, send_from_directory, request, Response, send_file
import openpyxl

app = Flask(__name__)

# Глобальный счётчик запусков (один для всех клиентов)
RUN_ID = 0

# Последний лог от клиента
CLIENT_LOG = {
    "run_id": None,
    "returncode": None,
    "stdout": "",
    "stderr": "",
}

# Результат проверки документа выгрузки: заголовки и строки, которые нельзя обработать
DOC_CHECK_RESULT = {"headers": [], "cannot_process_rows": []}

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ADMIN_HTML = """
<!doctype html>
<html lang="ru">
  <head>
    <meta charset="utf-8">
    <title>A.STORM робот — Админка</title>
    <style>
      * { box-sizing: border-box; }
      html, body { height: 100%; margin: 0; overflow: hidden; font-family: sans-serif; }
      .screen {
        display: none;
        position: fixed;
        top: 0; left: 0; right: 0; bottom: 0;
        overflow: hidden;
        flex-direction: column;
      }
      .screen.active {
        display: flex;
      }
      #home {
        align-items: center;
        justify-content: center;
        text-align: center;
        padding: 20px;
      }
      #home .main-title { margin-bottom: 30px; }
      #home .nav-btn { text-align: center; }
      .nav-btn {
        display: block;
        width: 100%;
        max-width: 320px;
        font-size: 16px;
        padding: 14px 20px;
        margin-bottom: 10px;
        cursor: pointer;
        text-align: left;
        border: 1px solid #333;
        border-radius: 6px;
        background: #f5f5f5;
      }
      .nav-btn:hover { background: #e8e8e8; }
      .back-btn {
        font-size: 14px;
        padding: 8px 16px;
        margin-bottom: 20px;
        cursor: pointer;
        background: #ddd;
        border: 1px solid #999;
        border-radius: 4px;
        flex-shrink: 0;
      }
      .back-btn:hover { background: #ccc; }
      .panel-content {
        flex: 1;
        min-height: 0;
        overflow: auto;
        padding: 0 20px 20px;
      }
      #status { margin-top: 15px; }
      #log { background: #111; color: #0f0; padding: 10px; white-space: pre-wrap; overflow: auto; }
      .main-title { margin-bottom: 30px; font-size: 1.5em; }
      .panel-header { flex-shrink: 0; padding: 20px 20px 0; }
      .doc-upload-area { margin: 15px 0; }
      .doc-upload-area input[type="file"] { margin-right: 10px; }
      .doc-upload-area button { padding: 8px 16px; cursor: pointer; }
      .doc-status { margin: 10px 0; min-height: 1.2em; }
      .doc-counts { margin-top: 20px; }
      .doc-counts button { margin-top: 15px; padding: 10px 16px; cursor: pointer; }
    </style>
  </head>
  <body>
    <!-- Главный экран -->
    <div id="home" class="screen active">
      <h1 class="main-title">A.STORM робот</h1>
      <button type="button" class="nav-btn" data-panel="remote">Дистанционный запуск</button>
      <button type="button" class="nav-btn" data-panel="doc">Проверить документ выгрузки</button>
      <button type="button" class="nav-btn" data-panel="check1c">Проверить данные 1С</button>
      <button type="button" class="nav-btn" data-panel="fill1c">Заполнить данные 1С</button>
      <button type="button" class="nav-btn" data-panel="load1c">Загрузить информацию из 1С</button>
    </div>

    <!-- Панель: Дистанционный запуск -->
    <div id="panel-remote" class="screen">
      <div class="panel-header">
        <button type="button" class="back-btn" data-back>Вернуться</button>
        <h1>Запуск alg.py на клиентах</h1>
        <p>Нажми кнопку, клиенты скачают свежий alg.py с сервера и выполнят его.</p>
        <button type="button" id="runBtn">Запустить на клиентах</button>
        <div id="status"></div>
      </div>
      <div class="panel-content">
        <h2>Лог клиента</h2>
        <pre id="log">Лога пока нет.</pre>
      </div>
    </div>

    <!-- Панель: Проверить документ выгрузки -->
    <div id="panel-doc" class="screen">
      <div class="panel-header">
        <button type="button" class="back-btn" data-back>Вернуться</button>
      </div>
      <div class="panel-content">
        <h1>Проверить документ выгрузки</h1>
        <p>Загрузите xlsx-документ, затем нажмите «Проверить». После проверки отобразится количество дел, которые можно и нельзя обработать.</p>
        <div class="doc-upload-area">
          <input type="file" id="docFile" accept=".xlsx,.xls" />
          <button type="button" id="docCheckBtn">Проверить</button>
        </div>
        <div id="docCheckStatus" class="doc-status"></div>
        <div id="docCounts" class="doc-counts" style="display:none;">
          <p><strong>Можно обработать:</strong> <span id="docCanProcess">0</span></p>
          <p><strong>Нельзя обработать:</strong> <span id="docCannotProcess">0</span></p>
          <button type="button" id="docDownloadBtn">Скачать документ с делами, которые нельзя обработать</button>
        </div>
      </div>
    </div>
    <div id="panel-check1c" class="screen">
      <div class="panel-header">
        <button type="button" class="back-btn" data-back>Вернуться</button>
      </div>
      <div class="panel-content">
        <h1>Проверить данные 1С</h1>
        <p>Раздел в разработке.</p>
      </div>
    </div>
    <div id="panel-fill1c" class="screen">
      <div class="panel-header">
        <button type="button" class="back-btn" data-back>Вернуться</button>
      </div>
      <div class="panel-content">
        <h1>Заполнить данные 1С</h1>
        <p>Раздел в разработке.</p>
      </div>
    </div>
    <div id="panel-load1c" class="screen">
      <div class="panel-header">
        <button type="button" class="back-btn" data-back>Вернуться</button>
      </div>
      <div class="panel-content">
        <h1>Загрузить информацию из 1С</h1>
        <p>Раздел в разработке.</p>
      </div>
    </div>

    <script>
      function show(id) {
        document.querySelectorAll('.screen').forEach(function(s) { s.classList.remove('active'); });
        var el = document.getElementById(id);
        if (el) el.classList.add('active');
      }

      document.addEventListener('DOMContentLoaded', function() {
        document.querySelectorAll('.nav-btn[data-panel]').forEach(function(btn) {
          btn.addEventListener('click', function(e) {
            e.preventDefault();
            var panelId = this.getAttribute('data-panel');
            if (panelId) show('panel-' + panelId);
          });
        });
        document.querySelectorAll('.back-btn[data-back]').forEach(function(btn) {
          btn.addEventListener('click', function() { show('home'); });
        });

        var docFile = document.getElementById('docFile');
        var docCheckBtn = document.getElementById('docCheckBtn');
        var docCheckStatus = document.getElementById('docCheckStatus');
        var docCounts = document.getElementById('docCounts');
        var docDownloadBtn = document.getElementById('docDownloadBtn');
        if (docCheckBtn && docFile) {
          docCheckBtn.addEventListener('click', function() {
            if (!docFile.files || !docFile.files[0]) {
              docCheckStatus.textContent = 'Сначала выберите файл .xlsx';
              return;
            }
            docCheckStatus.textContent = 'Проверяю...';
            var fd = new FormData();
            fd.append('document', docFile.files[0]);
            fetch('/check-doc-upload', { method: 'POST', body: fd })
              .then(function(r) { return r.json().then(function(d) { return { ok: r.ok, data: d }; }); })
              .then(function(o) {
                if (o.ok && o.data.ok) {
                  docCheckStatus.textContent = 'Проверка завершена.';
                  document.getElementById('docCanProcess').textContent = o.data.can_process;
                  document.getElementById('docCannotProcess').textContent = o.data.cannot_process;
                  docCounts.style.display = 'block';
                } else {
                  docCheckStatus.textContent = o.data.error || 'Ошибка проверки';
                }
              })
              .catch(function(e) {
                docCheckStatus.textContent = 'Ошибка: ' + e;
              });
          });
        }
        if (docDownloadBtn) {
          docDownloadBtn.addEventListener('click', function() {
            window.location.href = '/download-unprocessable-doc';
          });
        }
      });

      var runBtn = document.getElementById('runBtn');
      const statusEl = document.getElementById('status');
      const logEl = document.getElementById('log');

      async function loadLog() {
        try {
          const resp = await fetch('/client-log');
          const data = await resp.json();
          if (data.run_id === null) {
            logEl.textContent = 'Лога пока нет.';
          } else {
            logEl.textContent =
              'RUN_ID: ' + data.run_id + '\\n' +
              'returncode: ' + data.returncode + '\\n\\n' +
              'STDOUT:\\n' + (data.stdout || '') + '\\n\\n' +
              'STDERR:\\n' + (data.stderr || '');
          }
        } catch (e) {
          logEl.textContent = 'Ошибка загрузки лога: ' + e;
        }
      }

      if (runBtn) {
        runBtn.onclick = async () => {
          statusEl.textContent = 'Отправляю команду...';
          try {
            const resp = await fetch('/trigger', { method: 'POST' });
            const data = await resp.json();
            statusEl.textContent = 'OK, RUN_ID = ' + data.run_id;
          } catch (e) {
            statusEl.textContent = 'Ошибка: ' + e;
          }
        };
      }
      loadLog();
      setInterval(loadLog, 2000);
    </script>
  </body>
  </html>
"""


@app.route("/")
def admin() -> Response:
    # Простая админка с кнопкой
    return Response(ADMIN_HTML, mimetype="text/html")


@app.route("/trigger", methods=["POST"])
def trigger():
    """
    Админка дергает этот эндпоинт.
    Увеличиваем RUN_ID — клиенты по опросу увидят, что появился новый запуск.
    """
    global RUN_ID
    RUN_ID += 1
    return jsonify({"ok": True, "run_id": RUN_ID})


@app.route("/next-command", methods=["GET"])
def next_command():
    """
    Клиент опрашивает этот эндпоинт, передавая last_run_id (что он уже выполнял).
    Возвращаем, нужно ли запускать новый алгоритм, и ссылку на alg.py.
    """
    global RUN_ID
    try:
        last_run_id = int(request.args.get("last_run_id", "0"))
    except ValueError:
        last_run_id = 0

    should_run = RUN_ID > last_run_id
    alg_url = request.url_root.rstrip("/") + "/alg.py"

    return jsonify(
        {
            "run_id": RUN_ID,
            "should_run": should_run,
            "alg_url": alg_url,
        }
    )


@app.route("/alg.py", methods=["GET"])
def download_alg():
    """
    Клиент скачивает сам файл alg.py с сервера.
    """
    return send_from_directory(BASE_DIR, "alg.py", as_attachment=False)


@app.route("/client-log", methods=["POST", "GET"])
def client_log():
    """
    POST: клиент отправляет лог выполнения alg.py.
    GET: админка забирает последний лог.
    """
    global CLIENT_LOG
    if request.method == "POST":
        data = request.get_json(silent=True) or {}
        CLIENT_LOG = {
            "run_id": data.get("run_id"),
            "returncode": data.get("returncode"),
            "stdout": data.get("stdout") or "",
            "stderr": data.get("stderr") or "",
        }
        return jsonify({"ok": True})
    return jsonify(CLIENT_LOG)


def _parse_xlsx_rows(file_stream):
    """Читает xlsx из file-like объекта, возвращает (headers, list_of_row_dicts)."""
    wb = openpyxl.load_workbook(file_stream, read_only=True, data_only=True)
    sheet = wb.active
    rows = list(sheet.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return [], []
    headers = [str(h) if h is not None else "" for h in rows[0]]
    data = []
    for row in rows[1:]:
        data.append(dict(zip(headers, (v if v is not None else "" for v in row))))
    return headers, data


def _check_doc_placeholder(rows):
    """
    Заглушка проверки. Позже здесь будет реальная логика.
    Возвращает (can_process_count, cannot_process_rows).
    """
    # Пока считаем, что все можно обработать; необрабатываемых нет
    cannot_rows = []
    return len(rows) - len(cannot_rows), cannot_rows


@app.route("/check-doc-upload", methods=["POST"])
def check_doc_upload():
    """
    Принимает xlsx-файл, проверяет его, возвращает кол-во обрабатываемых и необрабатываемых.
    Сохраняет необрабатываемые строки для скачивания.
    """
    global DOC_CHECK_RESULT
    if "document" not in request.files:
        return jsonify({"error": "Файл не выбран"}), 400
    file = request.files["document"]
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "Нужен файл .xlsx"}), 400
    try:
        headers, rows = _parse_xlsx_rows(file.stream)
    except Exception as e:
        return jsonify({"error": f"Ошибка чтения файла: {e}"}), 400
    if not headers:
        return jsonify({"error": "Файл пустой или без заголовков"}), 400
    can_process, cannot_rows = _check_doc_placeholder(rows)
    DOC_CHECK_RESULT = {"headers": headers, "cannot_process_rows": cannot_rows}
    return jsonify({
        "ok": True,
        "can_process": can_process,
        "cannot_process": len(cannot_rows),
    })


@app.route("/download-unprocessable-doc", methods=["GET"])
def download_unprocessable_doc():
    """Отдаёт xlsx с делами, которые нельзя обработать (из последней проверки)."""
    global DOC_CHECK_RESULT
    headers = DOC_CHECK_RESULT.get("headers") or []
    rows = DOC_CHECK_RESULT.get("cannot_process_rows") or []
    if not headers and not rows:
        return jsonify({"error": "Нет данных для скачивания. Сначала выполните проверку документа."}), 404
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Необрабатываемые"
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)
    for r, row_dict in enumerate(rows, 2):
        for c, h in enumerate(headers, 1):
            ws.cell(row=r, column=c, value=row_dict.get(h, ""))
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="необрабатываемые_дела.xlsx",
    )


@app.route("/health")
def health():
    return jsonify({"status": "ok", "run_id": RUN_ID})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
