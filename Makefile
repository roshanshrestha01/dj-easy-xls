build:
	python3 -m build
deployPypiTest:
	python3 -m twine upload --repository testpypi dist/*
deploy:
	python3 -m twine upload dist/*