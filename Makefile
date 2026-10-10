PREFIX ?= /usr/local
BINDIR ?= $(PREFIX)/bin
LIBDIR ?= $(PREFIX)/lib/exceltool
INSTALL ?= install

MODULES := __init__.py __main__.py cli.py daemon.py daemon_client.py daemon_protocol.py daemon_runtime.py editing.py engine.py errors.py fonts.py locking.py operations.py output.py patching.py process_guard.py process_management.py ranges.py safety.py windows_job.py
INSTALLED_MODULES := $(addprefix $(LIBDIR)/,$(MODULES))

.PHONY: help install uninstall check

help:
	@echo "make install    安装到 $(PREFIX)（默认需要 sudo）"
	@echo "make uninstall  从 $(PREFIX) 卸载（默认需要 sudo）"
	@echo "make check      检查语法、入口和测试"
	@echo "可用 PREFIX=/path 覆盖安装前缀"

install:
	$(INSTALL) -d "$(BINDIR)" "$(LIBDIR)"
	$(INSTALL) -m 755 exceltool "$(BINDIR)/exceltool"
	$(INSTALL) -m 644 $(addprefix src/exceltool/,$(MODULES)) "$(LIBDIR)/"
	@echo "已安装: $(BINDIR)/exceltool"

uninstall:
	rm -f -- "$(BINDIR)/exceltool" $(addprefix ",$(addsuffix ",$(INSTALLED_MODULES)))
	-rmdir -- "$(LIBDIR)"
	@echo "已卸载: $(BINDIR)/exceltool"

check:
	python3 -m compileall -q src tests
	PYTHONPATH=src python3 -m unittest discover -s tests -v
	PYTHONPATH=src python3 -m exceltool --help >/dev/null
