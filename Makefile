# -*- makefile -*-
default: run

run:
	@echo "Runnin Purple Panda"
	-python PurplePanda.py

clean:
	-rm *~
	-rm *pyc
	-rm *hl7
