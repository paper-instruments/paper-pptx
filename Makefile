BEHAVE = behave

.PHONY: help
help:
	@echo "Please use \`make <target>' where <target> is one or more of"
	@echo "  accept       run acceptance tests using behave"
	@echo "  build        generate both sdist and wheel suitable for upload to PyPI"
	@echo "  clean        delete intermediate work product and start fresh"
	@echo "  coverage     run pytest with coverage"

.PHONY: accept
accept:
	$(BEHAVE) --stop

.PHONY: build
build:
	rm -rf dist
	python -m build
	twine check dist/*

.PHONY: clean
clean:
	find . -type f -name \*.pyc -exec rm {} \;
	find . -type f -name .DS_Store -exec rm {} \;
	rm -rf dist .coverage

.PHONY: coverage
coverage:
	py.test --cov-report term-missing --cov=pptx --cov=tests
