web: gunicorn -b 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 120 app:app
worker: python printer_worker.py --poll-seconds 1 --max-attempts 3 --retry-delay 2
