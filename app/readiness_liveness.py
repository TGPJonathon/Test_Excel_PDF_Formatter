from flask_healthz import HealthError

def print_ok():
  print("Everything is working")

def liveness():
  try:
    print_ok()
  except Exception:
    raise HealthError("Can't connect to API")

def readiness():
  try:
    print_ok()
  except Exception:
    raise HealthError("Can't connect to API")