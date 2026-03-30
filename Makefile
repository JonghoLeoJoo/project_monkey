VENV := .venv
PYTHON := $(VENV)/bin/python
PIP := $(VENV)/bin/pip

.PHONY: build run clean

build:
	python3 -m venv $(VENV)
	$(PIP) install -r requirements.txt

run: build
	$(PYTHON) main.py $(filter-out $@,$(MAKECMDGOALS))

clean:
	rm -rf $(VENV)
	rm -rf models

%:
	@: