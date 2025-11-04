SHELL:=/bin/bash

all:
	@echo python module_extractor.py ./config.json

install-os-deps:
	sudo apt install python3 -y
	sudo apt install python3-pip -y
	sudo apt install python3-venv -y
	sudo apt install python-is-python3 -y
	sudo apt autoremove

python-env:
	mkdir -p .python-venv
	python3 -m venv .python-venv
	. .python-venv/bin/activate; \
		python3 -m pip install --quiet --upgrade pip; \
		pip install --quiet -r requirements.txt;
	@echo always execute . .python-venv/bin/activate

pip-requirements:
	pip-compile requirements.in > requirements.txt
	. .python-venv/bin/activate; \
		pip install --quiet --upgrade -r requirements.txt

clean:
	rm -rf tests/results

.PHONY: tests
tests: tests/pattern.xlsx
	python test_module.py

tests/pattern.xlsx:
	python pattern.py -o ./tests

accept_results:
	cp tests/results/* tests/expected
