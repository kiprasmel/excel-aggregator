#!/usr/bin/env python3

import os
from math import inf
from pathlib import Path
import pandas as pd
import csv
from dataclasses import dataclass
from typing import List, Tuple, Callable, Optional, Any
from datetime import datetime

from excel_to_csv import excel_to_csv

@dataclass
class Location:
	x: int
	y: int
	value: Any
	sheet: List[List[str]]
	prefix: str = ""
	suffix: str = ""

	def goRight(self):
		return self._moveUntilValue(1, 0)

	def goBelow(self):
		return self._moveUntilValue(0, 1)

	def goRightUntilExact(self, target):
		return self._moveUntil(1, 0, lambda val: val == target)

	def goRightUntilPrefix(self, prefix):
		return self._moveUntil(1, 0, lambda val: str(val).startswith(prefix))

	def goRightUntilLastContinuousValue(self):
		return self._moveUntilLastContinuousValue(1, 0)

	def goBelowUntilExact(self, target):
		return self._moveUntil(0, 1, lambda val: val == target)

	def goBelowUntilPrefix(self, prefix):
		return self._moveUntil(0, 1, lambda val: str(val).startswith(prefix))

	def goBelowUntilLastContinuousValue(self):
		return self._moveUntilLastContinuousValue(0, 1)

	def move(self, dx: int, dy: int, moves: int):
		new_x, new_y = self.x, self.y
		while moves > 0:
			moves -= 1
			new_x += dx
			new_y += dy
			if not self._within_bounds(new_x, new_y):
				return None

		value = self._get_cell_value(new_x, new_y)
		return Location(new_x, new_y, value, self.sheet)

	# move until a value is reached, while within bounds
	def _moveUntilValue(self, dx, dy, moves=inf):
		new_x, new_y = self.x, self.y
		while moves > 0:
			moves -= 1
			new_x += dx
			new_y += dy
			if not self._within_bounds(new_x, new_y):
				return None

			value = self._get_cell_value(new_x, new_y)
			if value != "":
				return Location(new_x, new_y, value, self.sheet)
		
		return None

	def _moveUntil(self, dx, dy, condition):
		current = self
		while True:
			current = current._moveUntilValue(dx, dy)
			if not current:
				return None
			if condition(current.value):
				return current

	def _moveUntilLastContinuousValue(self, dx, dy):
		current = self
		while True:
			nxt = current.move(dx, dy, 1)
			if not nxt:
				return current
			if nxt.value == "":
				return current
			current = nxt

	def _get_cell_value(self, x, y):
		if self._within_bounds(x, y):
			return self.sheet[y][x]
		return None

	def _within_bounds(self, x, y):
		if 0 <= y < len(self.sheet) and 0 <= x < len(self.sheet[y]):
			return True
		return False

# move deltas
# POS = (dx, dy)
DOWN  = (0, 1)
RIGHT = (1, 0)
LEFT  = (-1, 0)
UP    = (0, -1)

class Finder:
	def __init__(self, find_func):
		self.find_func = find_func

	def move(self, delta, moves):
		(dx, dy) = delta
		return self._chain(lambda loc: loc.move(dx, dy, moves))

	def goRightUntilValue(self):
		return self._chain(lambda loc: loc.goRight())

	def goBelowUntilValue(self):
		return self._chain(lambda loc: loc.goBelow())

	def goRightUntilExact(self, value):
		return self._chain(lambda loc: loc.goRightUntilExact(value))

	def goRightUntilPrefix(self, prefix):
		return self._chain(lambda loc: loc.goRightUntilPrefix(prefix))

	def goBelowUntilLastContinuousValue(self):
		return self._chain(lambda loc: loc.goBelowUntilLastContinuousValue())

	def goBelowUntilExact(self, value):
		return self._chain(lambda loc: loc.goBelowUntilExact(value))

	def goBelowUntilPrefix(self, prefix):
		return self._chain(lambda loc: loc.goBelowUntilPrefix(prefix))

	def goRightUntilLastContinuousValue(self):
		return self._chain(lambda loc: loc.goRightUntilLastContinuousValue())

	def getSuffix(self):
		return self._chain(lambda loc: Location(loc.x, loc.y, loc.suffix, loc.sheet))

	def modify(self, func: Callable[[Any], Any]):
		return self._chain(lambda loc: Location(loc.x, loc.y, func(loc.value), loc.sheet))

	def _chain(self, operation):
		def new_find_func(sheet):
			loc = self.find_func(sheet)
			if loc:
				return operation(loc)
			return None
		return Finder(new_find_func)

	def __call__(self, sheet):
		return self.find_func(sheet)

def findExact(value: str):
	def finder(sheet):
		for y, row in enumerate(sheet):
			for x, cell in enumerate(row):
				if cell == value:
					return Location(x, y, cell, sheet)
		return None
	return Finder(finder)

def findPrefix(prefix: str):
	def finder(sheet):
		for y, row in enumerate(sheet):
			for x, cell in enumerate(row):
				if isinstance(cell, str) and cell.startswith(prefix):
					suffix = cell[len(prefix):]
					return Location(x, y, cell, sheet, prefix=prefix, suffix=suffix)
		return None
	return Finder(finder)

def aggregate_csv_data(folder_path: str, parse_columns: List[Tuple[str, Callable]]):
	all_data = []
	folder_name = os.path.basename(os.path.normpath(folder_path))
	timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
	outdir = "out"
	Path(outdir).mkdir(parents=True, exist_ok=True)
	output_file_csv = os.path.join(outdir, f"aggregated-{folder_name}-{timestamp}.csv")

	for filename in os.listdir(folder_path):
		if filename.endswith('.csv'):
			file_path = os.path.join(folder_path, filename)
			
			with open(file_path, 'r', newline='', encoding='utf-8') as csvfile:
				reader = csv.reader(csvfile)
				sheet = list(reader)

			row_data = {'Filename': filename}
			for column_name, value_fn in parse_columns:
				location = value_fn(sheet)
				if location:
					row_data[column_name] = location.value

			all_data.append(row_data)

	df = pd.DataFrame(all_data)
	df.to_csv(output_file_csv, index=False)
	return output_file_csv

PMV_AMOUNT = 1.21

def remove_pvm(value):
	divided = float(value) / PMV_AMOUNT
	rounded = round(divided, 2)
	return rounded

parse_columns_data1 = [
	("SF NUMERIS", findPrefix("PVM SĄSKAITA FAKTŪRA (VAT INVOICE)").modify(lambda x: x.split(" ")[-1])),
	("DATA", findPrefix("Išrašymo data / Date:").getSuffix()),
	("KODAS", findPrefix("Pirkėjas / Buyer").goBelowUntilPrefix("Įmones kodas:").getSuffix()),
	("PVM KODAS", findPrefix("Pirkėjas / Buyer").goBelowUntilPrefix("PVM kodas:").getSuffix()),
	("VARDAS PAVARDĖ/ĮM. PAVADINIMAS", findExact("Pirkėjas / Buyer").goBelowUntilValue()),
	("KAINA BE PVM", findExact("Bendros sumos EUR").goBelowUntilExact("Suma be PVM / total amount:").goRightUntilValue()),
]

parse_columns_data2 = [
	("SERIJA", findPrefix("Serija ").modify(lambda x: x.strip().split(" ")[1])),
	("NR", findPrefix("Serija ").modify(lambda x: x.strip().split(" ")[-1])),
	("DATA", findPrefix("Serija ").goBelowUntilValue().modify(lambda x: x.strip())),
	("PIRKEJAS", findExact("Pirkėjas:").move(RIGHT, 1).goBelowUntilExact("(pavadinimas)").move(UP, 1)),
	("KODAS", findExact("Pirkėjas:").move(RIGHT, 1).goBelowUntilExact("(pirkėjo kodas)").move(UP, 1)),
	("PVM KODAS", findExact("Pirkėjas:").move(RIGHT, 1).goBelowUntilExact("(PVM mokėtojo kodas)").move(UP, 1)),
	("KAINA BE PVM", findExact("Suma Eur").goBelowUntilLastContinuousValue().modify(remove_pvm)),
]

# select which parser to use
parse_columns = parse_columns_data2
excel_inputdir = "NS24 2024 08"

def main():
	# excel_inputdir = input("Enter the folder path containing excel files: ")
	excel_inputdir_name = os.path.basename(excel_inputdir)
	csv_outdir = os.path.join(excel_inputdir, f"csv-{excel_inputdir_name}")

	print(f"saving CSVs to '{csv_outdir}'")
	excel_to_csv(excel_inputdir, csv_outdir)

	print(f"aggregating data...")
	output_file_csv = aggregate_csv_data(csv_outdir, parse_columns)
	print(f"data aggregated to: {output_file_csv}")

if __name__ == "__main__":
	main()
