# EMClaw — developer entry points.
# The security demo needs only Python 3.9+ (no external services).
# The test/typecheck targets build the Microsoft 365 connector plugin.

CONNECTOR := v2/plugins/emclaw-m365-connector

.PHONY: help demo test typecheck install clean

help:
	@echo "EMClaw — available targets:"
	@echo "  make demo       Run the security-model walkthrough (Python stdlib only)"
	@echo "  make test       Build + unit-test the M365 connector plugin (Node)"
	@echo "  make typecheck  Type-check the M365 connector plugin (tsc --noEmit)"
	@echo "  make install    Install connector plugin dependencies (npm ci)"
	@echo "  make clean      Remove build output"

demo:
	@python3 demo/security_demo.py

install:
	@cd $(CONNECTOR) && npm ci

test: install
	@cd $(CONNECTOR) && npm run typecheck && npm run build && npm test

typecheck: install
	@cd $(CONNECTOR) && npx tsc -p tsconfig.json --noEmit && echo "typecheck: OK"

clean:
	@rm -rf $(CONNECTOR)/dist
