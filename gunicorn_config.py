# Gunicorn configuration file
bind = "0.0.0.0:5005"
workers = 4
worker_class = "sync"
timeout = 300  # 5 minutes timeout
keepalive = 5
max_requests = 1000
max_requests_jitter = 50
worker_tmp_dir = "/dev/shm"
preload_app = True
accesslog = "-"
errorlog = "-"
loglevel = "info" 