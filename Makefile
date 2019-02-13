all: csv

clean:
	rm csv/*.csv

csv:
	python ceus.py
