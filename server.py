"""
Сервер на Flask: общение по HTTP.
Нужно:
- alg.py лежит на сервере;
- ты в веб-админке жмёшь кнопку;
- клиенты, которые опрашивают сервер, видят новый запуск, СКАЧИВАЮТ alg.py и запускают его локально.

Запуск сервера: .venv/bin/python server.py
"""

import os
from flask import Flask, jsonify, send_from_directory, request, Response

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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ADMIN_HTML = """
<!doctype html>
<html lang="ru">
  <head>
    <meta charset="utf-8">
    <title>Админка запуска alg.py</title>
    <style>
      body { font-family: sans-serif; padding: 20px; }
      button { font-size: 16px; padding: 10px 20px; cursor: pointer; }
      #status { margin-top: 15px; }
      #log { background: #111; color: #0f0; padding: 10px; white-space: pre-wrap; max-height: 400px; overflow: auto; }
    </style>
  </head>
  <body>
    <h1>Запуск alg.py на клиентах</h1>
    <p>Нажми кнопку, клиенты скачают свежий alg.py с сервера и выполнят его.</p>
    <button id="runBtn">Запустить на клиентах</button>
    <div id="status"></div>
    <div style="margin-top:25px;">
      <h2>Лог клиента</h2>
      <pre id="log">Лога пока нет.</pre>
    </div>
    <script>
      const btn = document.getElementById('runBtn');
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

      btn.onclick = async () => {
        statusEl.textContent = 'Отправляю команду...';
        try {
          const resp = await fetch('/trigger', { method: 'POST' });
          const data = await resp.json();
          statusEl.textContent = 'OK, RUN_ID = ' + data.run_id;
        } catch (e) {
          statusEl.textContent = 'Ошибка: ' + e;
        }
      };

      loadLog();
      // автообновление лога раз в 2 секунды
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


@app.route("/health")
def health():
    return jsonify({"status": "ok", "run_id": RUN_ID})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
