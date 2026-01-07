PROJECT_NAME := $(notdir $(CURDIR))

all:
	python3 scripts/export_fab_outputs.py

release: all
	cd output/Gerbers && zip -r ../$(PROJECT_NAME)_Gerbers.zip .