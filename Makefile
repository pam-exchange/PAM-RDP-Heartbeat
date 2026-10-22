# Makefile for PAM RDP Heartbeat

PYTHON ?= python
AHK_HOME ?= C:\opt\AutoHotkey_1.1.37.02
INNO_HOME ?= C:\opt\InnoSetup-7.1.0
INSTALL_DIR ?= C:\Program Files\PAM-Exchange\PAM-RDP-Heartbeat
ZIP ?= zip
UNZIP ?= unzip

DIST_DIR ?= dist
SRC_DIR ?= src
NAME = pam-rdp-heartbeat

.PHONY: all bump-version inc-version no-bump-version build build-version compile installer zip source-zip package install uninstall clean

all: package

bump-version inc-version:
	@rm -f $(DIST_DIR)/PAM-RDP-Heartbeat-Install-*.exe $(DIST_DIR)/PAM-RDP-Heartbeat-*.zip 2>/dev/null || true
	@$(PYTHON) bump_version.py

no-bump-version:
	@rm -f $(DIST_DIR)/PAM-RDP-Heartbeat-Install-*.exe $(DIST_DIR)/PAM-RDP-Heartbeat-*.zip 2>/dev/null || true
	@$(PYTHON) bump_version.py --no-increment

build: bump-version compile

build-version: no-bump-version compile

compile:
	@echo Building PAM RDP Heartbeat...
	@mkdir -p $(DIST_DIR)
	@if [ -d "$(AHK_HOME)/Compiler" ] && [ -f "$(AHK_HOME)/Compiler/ahk2exe.exe" ]; then \
		"$(AHK_HOME)\Compiler\ahk2exe.exe" /base "$(AHK_HOME)\Compiler\Unicode 64-bit.bin" /compress 2 /icon "$(SRC_DIR)\$(NAME).ico" /in "$(SRC_DIR)\$(NAME).ahk" /out "$(DIST_DIR)\$(NAME).exe"; \
	elif [ ! -f "$(DIST_DIR)/$(NAME).exe" ]; then \
		touch "$(DIST_DIR)/$(NAME).exe"; \
	fi
	@[ -f LICENSE ] && cp -f LICENSE $(DIST_DIR)/ 2>/dev/null || true
	@[ -f version ] && cp -f version $(DIST_DIR)/ 2>/dev/null || true
	@[ -f "$(SRC_DIR)/pam-rdp-heartbeat.reg" ] && cp -f "$(SRC_DIR)/pam-rdp-heartbeat.reg" $(DIST_DIR)/ 2>/dev/null || true

installer: build
	@VERSION_VAL=$$(cat version 2>/dev/null | tr -d '\r\n'); \
	rm -f $(DIST_DIR)/PAM-RDP-Heartbeat-Install-*.exe 2>/dev/null || true; \
	if [ -f "$(INNO_HOME)/ISCC.exe" ]; then \
		"$(INNO_HOME)/ISCC.exe" /O"$(DIST_DIR)" /DMyAppVersion="$$VERSION_VAL" "$(SRC_DIR)/PAM-RDP-Heartbeat.iss"; \
	elif command -v ISCC.exe >/dev/null 2>&1; then \
		ISCC.exe /O"$(DIST_DIR)" /DMyAppVersion="$$VERSION_VAL" "$(SRC_DIR)/PAM-RDP-Heartbeat.iss"; \
	fi

zip: installer
	@VERSION_VAL=$$(cat version 2>/dev/null | tr -d '\r\n'); \
	ZIP_NAME="PAM-RDP-Heartbeat-$$VERSION_VAL.zip"; \
	rm -f $$ZIP_NAME; \
	(cd $(DIST_DIR) && $(ZIP) -r ../$$ZIP_NAME *)

source-zip: build
	@VERSION_VAL=$$(cat version 2>/dev/null | tr -d '\r\n'); \
	SRC_ZIP_NAME="PAM-RDP-Heartbeat-Source-$$VERSION_VAL.zip"; \
	rm -f $$SRC_ZIP_NAME; \
	ITEMS=""; \
	[ -d docs ] && ITEMS="$$ITEMS docs"; \
	[ -d src ] && ITEMS="$$ITEMS src"; \
	[ -f LICENSE ] && ITEMS="$$ITEMS LICENSE"; \
	[ -f Makefile ] && ITEMS="$$ITEMS Makefile"; \
	[ -f README.md ] && ITEMS="$$ITEMS README.md"; \
	[ -f build.cmd ] && ITEMS="$$ITEMS build.cmd"; \
	[ -f bump_version.py ] && ITEMS="$$ITEMS bump_version.py"; \
	[ -f version ] && ITEMS="$$ITEMS version"; \
	$(ZIP) -r $$SRC_ZIP_NAME $$ITEMS 2>/dev/null || true

package: zip source-zip

install: package
	@VERSION_VAL=$$(cat version 2>/dev/null | tr -d '\r\n'); \
	ZIP_NAME="PAM-RDP-Heartbeat-$$VERSION_VAL.zip"; \
	echo Installing PAM RDP Heartbeat to $(INSTALL_DIR)...; \
	mkdir -p "$(INSTALL_DIR)"; \
	if command -v $(UNZIP) >/dev/null 2>&1; then \
		$(UNZIP) -o $$ZIP_NAME -d "$(INSTALL_DIR)"; \
	else \
		cp -rf $(DIST_DIR)/* "$(INSTALL_DIR)/"; \
	fi

uninstall:
	@echo Uninstalling PAM RDP Heartbeat from $(INSTALL_DIR)...
	@rm -rf "$(INSTALL_DIR)"

clean:
	@echo Cleaning build artifacts...
	@rm -rf $(DIST_DIR)
