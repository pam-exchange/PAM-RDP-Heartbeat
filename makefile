# Makefile for PAM RDP Heartbeat

AHK_HOME ?= C:\opt\AutoHotkey_1.1.37.02
INNO_HOME ?= C:\opt\InnoSetup-7.1.0
INSTALL_DIR ?= C:\Program Files\PAM-Exchange\PAM-RDP-Heartbeat

DIST = dist
SRC = src
NAME = pam-rdp-heartbeat

PYTHON ?= python3

VERSION = $(shell cat version 2>/dev/null || echo 1.1.0.0)

ifeq ($(OS),Windows_NT)
    RM_DIST = if exist $(DIST) rmdir /S /Q $(DIST)
    RM_ZIPS = if exist PAM-RDP-Heartbeat-*.zip del /Q /F PAM-RDP-Heartbeat-*.zip
    MKDIR_DIST = if not exist $(DIST) mkdir $(DIST)
    CP = copy /Y
    ZIP_CMD = zip -r
    RM_INSTALL = if exist "$(INSTALL_DIR)" rmdir /S /Q "$(INSTALL_DIR)"
    AHK_EXE = "$(AHK_HOME)\Compiler\ahk2exe.exe"
    AHK_BIN = "$(AHK_HOME)\Compiler\Unicode 64-bit.bin"
    ISCC_EXE = "$(INNO_HOME)\ISCC.exe"
else
    RM_DIST = rm -rf $(DIST)
    RM_ZIPS = rm -f PAM-RDP-Heartbeat-*.zip
    MKDIR_DIST = mkdir -p $(DIST)
    CP = cp
    ZIP_CMD = zip -r
    RM_INSTALL = rm -rf "$(INSTALL_DIR)"
    AHK_EXE = "$(AHK_HOME)/Compiler/ahk2exe.exe"
    AHK_BIN = "$(AHK_HOME)/Compiler/Unicode 64-bit.bin"
    ISCC_EXE = "$(INNO_HOME)/ISCC.exe"
endif

.PHONY: all build package clean install uninstall inc-version

all: build package

inc-version:
	@$(PYTHON) scripts/update_version.py || python scripts/update_version.py

build: inc-version
	$(eval VERSION := $(shell cat version))
	@echo "Building PAM RDP Heartbeat version $(VERSION)..."
	@$(MKDIR_DIST)
	-$(AHK_EXE) /base $(AHK_BIN) /compress 2 /icon "$(SRC)/$(NAME).ico" /in "$(SRC)/$(NAME).ahk" /out "$(DIST)/$(NAME).exe"
	-$(CP) LICENSE $(DIST)
	-$(CP) "$(SRC)/pam-rdp-heartbeat.reg" $(DIST)
	-$(ISCC_EXE) /O"$(DIST)" /DMyAppVersion="$(VERSION)" "$(SRC)/PAM-RDP-Heartbeat.iss"

package:
	$(eval VERSION := $(shell cat version))
	@echo "Packaging PAM RDP Heartbeat version $(VERSION)..."
	@$(RM_ZIPS)
	cd $(DIST) && $(ZIP_CMD) ../PAM-RDP-Heartbeat-$(VERSION).zip *
	$(ZIP_CMD) PAM-RDP-Heartbeat-Source-$(VERSION).zip docs src scripts build.cmd makefile README.md LICENSE version

install:
	$(eval VERSION := $(shell cat version))
	@echo "Installing PAM RDP Heartbeat version $(VERSION) to $(INSTALL_DIR)..."
ifeq ($(OS),Windows_NT)
	powershell -Command "if (!(Test-Path 'PAM-RDP-Heartbeat-$(VERSION).zip')) { Write-Error 'Package PAM-RDP-Heartbeat-$(VERSION).zip not found. Run make package first.'; exit 1 }; if (!(Test-Path '$(INSTALL_DIR)')) { New-Item -ItemType Directory -Force -Path '$(INSTALL_DIR)' }; Expand-Archive -Force -Path 'PAM-RDP-Heartbeat-$(VERSION).zip' -DestinationPath '$(INSTALL_DIR)'"
else
	@if [ ! -f "PAM-RDP-Heartbeat-$(VERSION).zip" ]; then \
		echo "Package PAM-RDP-Heartbeat-$(VERSION).zip not found. Run 'make package' first."; \
		exit 1; \
	fi
	mkdir -p "$(INSTALL_DIR)"
	unzip -o PAM-RDP-Heartbeat-$(VERSION).zip -d "$(INSTALL_DIR)"
endif

uninstall:
	@echo "Uninstalling PAM RDP Heartbeat from $(INSTALL_DIR)..."
	@$(RM_INSTALL)

clean:
	@echo "Cleaning build artifacts..."
	-@$(RM_DIST)
	-@$(RM_ZIPS)
