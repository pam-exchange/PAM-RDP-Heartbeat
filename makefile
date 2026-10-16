# Makefile for PAM RDP Heartbeat

SHELL = cmd.exe

INNO_HOME ?= C:\opt\InnoSetup-7.1.0
INSTALL_DIR ?= C:\Program Files\PAM-Exchange\PAM-RDP-Heartbeat

DIST = dist
SRC = src
NAME = pam-rdp-heartbeat

PYTHON ?= python

.PHONY: all build package clean install uninstall inc-version

all: build package

inc-version:
	@$(PYTHON) scripts/update_version.py

build: inc-version
	@echo Building PAM RDP Heartbeat...
	-if not exist $(DIST) mkdir $(DIST)
	-if exist $(DIST)\PAM-RDP-Heartbeat-Install-*.exe del /Q /F $(DIST)\PAM-RDP-Heartbeat-Install-*.exe
	-if exist $(DIST)\PAM-RDP-Heartbeat.exe del /Q /F $(DIST)\PAM-RDP-Heartbeat.exe
	-$(PYTHON) -m PyInstaller --noconsole --onefile --icon="$(SRC)\$(NAME).ico" --name="$(NAME)" --distpath="$(DIST)" --workpath="build" "$(SRC)\$(NAME).py"
	-copy /Y LICENSE $(DIST)
	-copy /Y version $(DIST)
	-copy /Y "$(SRC)\pam-rdp-heartbeat.reg" $(DIST)
	-"$(INNO_HOME)\ISCC.exe" /O"$(DIST)" /DMyAppVersion="$(shell $(PYTHON) -c "print(open('version').read().strip())")" "$(SRC)\PAM-RDP-Heartbeat.iss"

package:
	@echo Packaging PAM RDP Heartbeat...
	cd $(DIST) && zip -r ..\PAM-RDP-Heartbeat-$(shell $(PYTHON) -c "print(open('version').read().strip())").zip *
	zip -r PAM-RDP-Heartbeat-Source-$(shell $(PYTHON) -c "print(open('version').read().strip())").zip docs src scripts build.cmd makefile README.md LICENSE version

install:
	@echo Installing PAM RDP Heartbeat to $(INSTALL_DIR)...
	powershell -Command "$$v = (Get-Content version).Trim(); if (!(Test-Path \"PAM-RDP-Heartbeat-$$v.zip\")) { Write-Error \"Package PAM-RDP-Heartbeat-$$v.zip not found. Run make package first.\"; exit 1 }; if (!(Test-Path '$(INSTALL_DIR)')) { New-Item -ItemType Directory -Force -Path '$(INSTALL_DIR)' }; Expand-Archive -Force -Path \"PAM-RDP-Heartbeat-$$v.zip\" -DestinationPath '$(INSTALL_DIR)'"

uninstall:
	@echo Uninstalling PAM RDP Heartbeat from $(INSTALL_DIR)...
	-if exist "$(INSTALL_DIR)" rmdir /S /Q "$(INSTALL_DIR)"

clean:
	@echo Cleaning build artifacts...
	-if exist $(DIST) rmdir /S /Q $(DIST)
	-if exist build rmdir /S /Q build
	-if exist $(NAME).spec del /Q /F $(NAME).spec
	-if exist PAM-RDP-Heartbeat-*.zip del /Q /F PAM-RDP-Heartbeat-*.zip
