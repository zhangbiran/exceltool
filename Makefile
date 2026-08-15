PREFIX ?= /usr/local
BINDIR ?= $(PREFIX)/bin
LIBDIR ?= $(PREFIX)/lib/exceltool
INSTALL ?= install
PLUGIN_EXCEL_SKILL ?=

MODULES := __init__.py __main__.py cli.py editing.py engine.py errors.py fonts.py operations.py output.py patching.py ranges.py safety.py
INSTALLED_MODULES := $(addprefix $(LIBDIR)/,$(MODULES))

.PHONY: help install uninstall check sync-skill-docs check-skill-docs

help:
	@echo "make install    安装到 $(PREFIX)（默认需要 sudo）"
	@echo "make uninstall  从 $(PREFIX) 卸载（默认需要 sudo）"
	@echo "make check      检查语法、入口和测试"
	@echo "make sync-skill-docs PLUGIN_EXCEL_SKILL=/path  同步插件参考文档"
	@echo "make check-skill-docs PLUGIN_EXCEL_SKILL=/path 检查插件文档一致性"
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

sync-skill-docs:
	@test -n "$(PLUGIN_EXCEL_SKILL)" || { echo "必须指定 PLUGIN_EXCEL_SKILL" >&2; exit 2; }
	$(INSTALL) -d "$(PLUGIN_EXCEL_SKILL)/references"
	$(INSTALL) -m 644 README.md "$(PLUGIN_EXCEL_SKILL)/references/cli.md"

check-skill-docs:
	@test -n "$(PLUGIN_EXCEL_SKILL)" || { echo "必须指定 PLUGIN_EXCEL_SKILL" >&2; exit 2; }
	cmp README.md "$(PLUGIN_EXCEL_SKILL)/references/cli.md"
