import sys
import os
import re
from pathlib import Path
from subprocess import call
from pdfminer.high_level import extract_text

# イケてないcsvから項目を抽出する
class Parser:
	def __init__(self, text):
		self.lines = text.split('\n')

	def find_index_with_string(self, target_string):
		for index, item in enumerate(self.lines):
				if target_string in item:
						return index
		return -1  # ターゲット文字列が見つからない場合

	# pdfから取り出したテキストからf_wordが含まれる行からt_wordの含まれる行までを取り出す
	# テキストがまともに区切られていないため、項目名による検索が不可能だったため仕方なく。
	def extract_from_to(self, f_word, t_word):
		f_index = self.find_index_with_string(f_word)
		if (f_index < 0):
			 return ""
		t_index = self.find_index_with_string(t_word)

		if (t_index < 0):
			 return ""

		extracted = self.lines[f_index:t_index]
		return '\n'.join(extracted)

	def extract_properties(self):
		company_name = next((l for l in self.lines if ('株式会社' in l)), None)

		# url
		url_line = next((l for l in self.lines if ('https://' in l)), None).replace(' ','')
		url = re.search(r'https://[\w.\/\-\&\?]+', url_line).group(0)

		job = next((l for l in self.lines if ('職種:' in l)), None).replace(' ','').replace('職種:', '')
		position = next((l for l in self.lines if ('職位:' in l)), None).replace(' ','').replace('職位:', '')
		salary = next((l for l in self.lines if ('想定年収:' in l)), None).replace(' ','').replace('想定年収:', '')

		condition = self.extract_from_to("内定の可能性が高い人", "書類⾒送りの主な理由")

		# location
		location_challenge_1 = self.extract_from_to("〈勤務地〉", "〈勤務時間〉")

		if location_challenge_1 != "〈勤務地〉\n":
			location = location_challenge_1.replace(' ','').replace('〈勤務地〉\n','')
		else:
			location_challenge_2 = self.extract_from_to("〈勤務時間〉", "〈補足情報〉")
			location = location_challenge_2.replace(' ','').replace('〈勤務時間〉\n','')

		return {"企業名": company_name, "職種": job, "職位": position, "想定年収": salary, "会社HP": url, "必要条件・内定の可能性が高い人": condition, "勤務地": location}

# file_pathにファイルがなければ作成
def create_file_if_not_exists(file_path):
	if not os.path.exists(file_path):
		with open(file_path, 'w') as file:
			pass

# Parseクラスでパースした項目をcsvに出力
def output_csv(properties_list):
	file_path = "./out.csv"
	create_file_if_not_exists(file_path)

	columns = ["企業名","会社HP","勤務地","職種","職位","想定年収","必要条件・内定の可能性が高い人"]

	with open(file_path, 'w') as file:
		file.write(",".join(columns) + "\n")
		for properties in properties_list:
			values = []
			for c in columns:
				values.append(f"\"{properties[c]}\"")
			file.write(f"{','.join(values)}\n")

# pdf2txt.py のパス
py_path = Path(sys.exec_prefix) / "Scripts" / "pdf2txt.py"

# ファイル名の取得
dir_path = "files"
files = os.listdir(dir_path)

# pdfファイルのみファイル名を取得
f_end = [s for s in files if s.endswith('.pdf')]

ls = []

for file in f_end:
	text = extract_text(f"{dir_path}/{file}")
	parser = Parser(text)
	properties = parser.extract_properties()
	ls.append(properties)

output_csv(ls)