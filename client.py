import sys
import time
import subprocess

try:
    import requests
except ImportError:
    print("Установите requests (в venv): pip install requests")
    sys.exit(1)


# Адрес сервера, где крутится Flask и лежит alg.py
SERVER_URL = "http://localhost:8080"  # поменяйте на http://IP_СЕРВЕРА:8080 при необходимости
POLL_INTERVAL = 2  # секунды между опросами /next-command


def download_and_run(alg_url: str, run_id: int) -> None:
    print(f"[*] RUN_ID={run_id}: качаю alg.py с {alg_url} ...")
    r = requests.get(alg_url, timeout=30)
    r.raise_for_status()

    local_name = "alg.py"
    with open(local_name, "wb") as f:
        f.write(r.content)

    print(f"[*] RUN_ID={run_id}: alg.py сохранён, запускаю...")
    # Выполняем скачанный alg.py локально и собираем вывод
    result = subprocess.run(
        [sys.executable, local_name],
        cwd=".",
        capture_output=True,
        text=True,
    )
    print(f"[*] RUN_ID={run_id}: alg.py завершился с кодом {result.returncode}.")

    # Отправляем лог выполнения на сервер, чтобы он отобразился в админке
    try:
        requests.post(
            f"{SERVER_URL}/client-log",
            json={
                "run_id": run_id,
                "returncode": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
            },
            timeout=15,
        )
    except requests.RequestException as e:
        print(f"[!] Не удалось отправить лог на сервер: {e}")


def main():
    last_run_id = 0
    print(f"Клиент опрашивает сервер {SERVER_URL}, интервал {POLL_INTERVAL} с.")

    while True:
        try:
            resp = requests.get(
                f"{SERVER_URL}/next-command",
                params={"last_run_id": last_run_id},
                timeout=15,
            )
            data = resp.json()
            run_id = int(data.get("run_id", 0))
            should_run = bool(data.get("should_run"))
            alg_url = data.get("alg_url")

            if should_run and alg_url and run_id > last_run_id:
                download_and_run(alg_url, run_id)
                last_run_id = run_id

        except requests.RequestException as e:
            print(f"[!] Ошибка HTTP-запроса: {e}")
        except Exception as e:
            print(f"[!] Ошибка клиента: {e}")

        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    main()
