import requests
r = requests.post("http://localhost:11434/api/generate", json={
    "model": "qwen2.5:7b-instruct",
    "prompt": 'Верни JSON {"ok": true}',
    "format": "json", "stream": False})
print(r.json()["response"])