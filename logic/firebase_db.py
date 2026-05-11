import json
import os

import firebase_admin
from firebase_admin import credentials, db as rtdb

_app = None

DATABASE_URL = 'https://picking-d3107-default-rtdb.firebaseio.com'


def _init():
    global _app
    if _app is not None:
        return True
    creds_json = os.environ.get('FIREBASE_CREDENTIALS')
    if not creds_json:
        return False
    try:
        cred_dict = json.loads(creds_json)
        cred = credentials.Certificate(cred_dict)
        _app = firebase_admin.initialize_app(cred, {'databaseURL': DATABASE_URL})
        return True
    except Exception:
        return False


def is_available():
    return bool(os.environ.get('FIREBASE_CREDENTIALS'))


def load_estructura():
    _init()
    data = rtdb.reference('/estructura').get()
    return data or {}


def save_estructura(data):
    _init()
    rtdb.reference('/estructura').set(data)


