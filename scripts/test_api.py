"""Quick test for the config API using TestClient."""
from fastapi.testclient import TestClient
from shiftortools.api import app, CFG_PATH
import json


def run_tests():
    client = TestClient(app)

    # Ensure config file is removed for clean test
    try:
        CFG_PATH.unlink()
    except Exception:
        pass

    # GET should return empty dict
    r = client.get('/api/config')
    assert r.status_code == 200
    assert r.json() == {}

    # PUT with invalid payload
    bad = {'大学病院': {'mon': 2}}
    r = client.put('/api/config', json=bad)
    assert r.status_code == 400

    # PUT valid payload
    good = {'大学病院': {"0": 2, "1": 2, "2": 2, "3": 2, "4": 2, "5": 0, "6": 0}}
    r = client.put('/api/config', json=good)
    assert r.status_code == 200
    assert r.json().get('status') == 'ok'

    r = client.get('/api/config')
    assert r.status_code == 200
    j = r.json()
    assert '大学病院' in j
    print('API tests passed')


if __name__ == '__main__':
    run_tests()
