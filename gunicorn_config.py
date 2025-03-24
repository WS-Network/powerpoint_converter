import multiprocessing

# Server socket
bind = "0.0.0.0:10000"
backlog = 2048

# Worker processes
workers = 2  # Reduced number of workers but with more resources per worker
worker_class = 'gevent'  # Use gevent for async workers - better for I/O bound operations
threads = 2  # Reduced threads to avoid memory fragmentation
worker_connections = 1000
timeout = 900  # 15 minutes - increased timeout for large files
keepalive = 2

# Logging
accesslog = '-'
errorlog = '-'
loglevel = 'info'
capture_output = True  # Capture stdout/stderr from workers
log_file = '-'  # Log to stdout

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
max_requests = 1000
max_requests_jitter = 50

# Worker performance tuning
worker_tmp_dir = '/dev/shm'  # Use shared memory for temporary files
forwarded_allow_ips = '*'  # Trust X-Forwarded-* headers
graceful_timeout = 90  # Give workers 90 seconds to finish processing requests

# Memory optimization
limit_request_line = 0  # No limit on the size of the HTTP request line
limit_request_fields = 100  # Maximum number of HTTP headers
limit_request_field_size = 0  # No limit on the size of the HTTP header 