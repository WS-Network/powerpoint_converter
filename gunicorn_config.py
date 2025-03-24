import multiprocessing

# Server socket
bind = "0.0.0.0:10000"
backlog = 2048

# Worker processes
workers = 2  # Fixed number of workers for better stability
worker_class = 'gevent'
threads = 2
worker_connections = 100
timeout = 300  # 5 minutes
keepalive = 2

# Logging
accesslog = '-'
errorlog = '-'
loglevel = 'info'

# Process naming
proc_name = 'slide_harmony'

# Server mechanics
daemon = False
pidfile = None
umask = 0
user = None
group = None
tmp_upload_dir = None

# Maximum requests per worker
max_requests = 100
max_requests_jitter = 10

# Memory management
max_worker_lifetime = 300  # 5 minutes
max_worker_lifetime_jitter = 10 