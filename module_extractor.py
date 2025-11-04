import os
import json
import csv
import sys
from worker import Worker

class ModuleExtractor:
    def __init__(self, config_file, output_dir=None):
        self.config_file = config_file
        self.output_dir = output_dir or os.getcwd()
        self.modules = []

        if not os.path.isfile(self.config_file):
            raise FileNotFoundError(f"Config file not found: {self.config_file}")
        if not os.path.isdir(self.output_dir):
            os.makedirs(self.output_dir, exist_ok=True)

        with open(self.config_file, 'r') as f:
            self.config = json.load(f)

    def load_config(self):
        with open(self.config_file, "r", encoding="utf-8") as f:
            return json.load(f)

    def run(self):
        if not self.output_dir:
            return
        base_name = os.path.splitext(os.path.basename(self.config_file))[0]
        dummy_file_path = os.path.join(self.output_dir, f"{base_name}.csv")

        config = self.load_config()
        worker = Worker(config)
        headers, rows = worker.extract()

        with open(dummy_file_path, 'w', newline='') as f:
            writer = csv.writer(f)
            # https://github.com/python/cpython/blob/main/Lib/csv.py
            # writer = csv.writer(
            #     f,
            #     delimiter=",",        # field separator: common are ',', ';', '\t', '|'
            #     quotechar='"',        # character used to quote fields containing special chars
            #     quoting=csv.QUOTE_MINIMAL,  # controls when quoting occurs
            #     escapechar="\\",      # used to escape delimiter or quotechar if quoting=QUOTE_NONE
            #     doublequote=True,     # if True, quotechar is doubled inside fields instead of escaped
            #     lineterminator="\n"   # line separator, usually '\n' or '\r\n'
            # )
            writer.writerow(headers)
            writer.writerows(rows)

def main():
    if len(sys.argv) < 2:
        print("Usage: python module_extractor.py <config_file> [output_dir]")
        sys.exit(1)

    config_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None

    extractor = ModuleExtractor(config_file, output_dir)
    extractor.run()

if __name__ == "__main__":
    main()
