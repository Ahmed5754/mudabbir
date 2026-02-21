# Mudabbir Windows Quickstart

## 1) Get latest code

```powershell
git pull
```

## 2) Install/update dependencies

If you use `uv`:

```powershell
uv sync
```

Or with pip:

```powershell
python -m pip install -U pip
pip install -e .
```

## 3) Run tests

Full test suite:

```powershell
pytest -q
```

Focused tests for Windows intent/fastpath:

```powershell
pytest -q tests/test_windows_intent_map.py tests/test_agent_loop_fastpath.py
```

## 4) Start Mudabbir

```powershell
mudabbir
```

If command is not found:

```powershell
python -m Mudabbir
```

## 4.1) If CTRL+C does not stop it

Use built-in force stop:

```powershell
mudabbir --stop
```

## 5) Quick smoke commands to try

- `control mouse`
- `control keyboard`
- `control folders`
- `control color`
- `control desktop`
- `netstat -b`
- `netsh wlan show profiles`
- `nbtstat -a FILESRV`
