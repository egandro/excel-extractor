import os
import json
import csv
import sys
import pandas as pd
from openpyxl.utils.cell import range_boundaries

class Worker:
    def __init__(self, cfg):
        self.cfg = cfg
        self.src = cfg.get("source")
        self.sheet_name = cfg.get("sheet")
        self.hdr_entries = cfg.get("headers", [])
        self.data_entries = cfg.get("data", [])
        self.sheet = None
        self.header_blocks = []
        self.rows = []
        self.final_header = []
        self.na = cfg.get("na", "")

        required = {
            "source": self.src,
            "sheet": self.sheet_name,
            "headers": self.hdr_entries,
        }

        for key, value in required.items():
            if not value:
                raise ValueError(f"Configuration key '{key}' is missing or empty")

    def _load_sheet(self):
        self.sheet = pd.read_excel(self.src, sheet_name=self.sheet_name, header=None)

    def _offset_by_selector(self, sel: str, idx: int, context: str, allow_multirow: bool, allow_single_cell: bool):
        if not isinstance(sel, str):
            raise ValueError(f"{context} {idx} / selector: Must be a string")

        if not sel or not sel.strip():
            raise ValueError(f"{context} {idx} / selector: Cannot be null or empty")

        sel = sel.strip()
        if ":" not in sel:
            raise ValueError(f"{context} {idx} / selector '{sel}': Must contain ':'")

        try:
            min_col, min_row, max_col, max_row = range_boundaries(sel)
        except Exception as e:
            raise ValueError(f"{context} {idx} / selector '{sel}': Invalid format") from e

        if min_row > max_row or min_col > max_col:
            raise ValueError(f"{context} {idx} / selector '{sel}': Invalid order")

        is_single_cell = (min_col == max_col and min_row == max_row)
        if is_single_cell:
            if not allow_single_cell:
                raise ValueError(f"{context} {idx} / selector '{sel}': Single cell not allowed")
            return min_col, min_row, max_col, max_row

        if not allow_multirow and min_row != max_row:
            raise ValueError(f"{context} {idx} / selector '{sel}': Must be single row")

        return min_col, min_row, max_col, max_row

    def _build_header_blocks(self):
        blocks = []
        for idx, entry in enumerate(self.hdr_entries, start=1):
            count = sum(k in entry for k in ("static", "fixed", "range"))
            if count != 1:
                raise ValueError(f"Header {idx}: Header entries must be static/fixed/range")

            if "static" in entry:
                blocks.append({"type": "static", "value": entry["static"].strip()})
            elif "fixed" in entry:
                sel = entry["fixed"]
                min_c, min_r, max_c, max_r = self._offset_by_selector(sel, idx, "Header", False, True)
                if min_c != max_c or min_r != max_r:
                    raise ValueError(f"Header fixed '{sel}' must be single cell")
                cell = self.sheet.iat[min_r - 1, min_c - 1]
                blocks.append({"type": "fixed", "value": str(cell) if pd.notna(cell) else ""})
            elif "range" in entry:
                sel = entry["range"]
                min_c, min_r, max_c, _ = self._offset_by_selector(sel, idx, "Header", False, False)
                values = [str(self.sheet.iat[min_r - 1, c - 1]) if pd.notna(self.sheet.iat[min_r - 1, c - 1]) else "" for c in range(min_c, max_c + 1)]
                blocks.append({"type": "range", "col_count": max_c - min_c + 1, "values": values})
        self.header_blocks = blocks

    def _extract_data_rows(self):
        data_rows = 0
        data_columns = 0
        data_rows_range = 0
        last_range_idx = -1
        last_range_row_count = -1
        data_blocks = []

        for idx, entry in enumerate(self.data_entries, start=1):
            count = sum(k in entry for k in ("static", "fixed", "range"))
            if count != 1:
                raise ValueError(f"Data {idx}: Data descriptions must be static/fixed/range")

            if "static" in entry:
                sel = entry["static"]
                data_blocks.append({"type": "static", "value": entry["static"].strip()})
                data_rows += 1
                data_columns += 1
            elif "fixed" in entry:
                sel = entry["fixed"]
                min_c, min_r, max_c, max_r = self._offset_by_selector(sel, idx, "Data", True, True)
                if min_c != max_c or min_r != max_r:
                    raise ValueError(f"Data fixed '{sel}' must be single cell")
                cell = self.sheet.iat[min_r - 1, min_c - 1]
                data_blocks.append({"type": "fixed", "value": str(cell) if pd.notna(cell) else self.na})
                data_rows += 1
                data_columns += 1
            elif "range" in entry:
                sel = entry["range"]
                min_c, min_r, max_c, max_r = self._offset_by_selector(sel, idx, "Data", True, False)
                col_count = max_c - min_c + 1
                row_count = max_r - min_r + 1
                if last_range_idx < 0:
                    last_range_idx = idx
                    last_range_row_count = row_count
                else:
                    if last_range_row_count != row_count:
                        print("last_range_row_count", last_range_row_count)
                        print("row_count", row_count)
                        raise ValueError(f"Data {idx}: Data rows count differs to Data {last_range_idx}")
                data_blocks.append({
                    "type": "range",
                    "col_offset": data_columns,
                    "col_count": col_count,
                    "row_count": row_count,
                    "min_r": min_r, "max_r": max_r,
                    "min_c": min_c, "max_c": max_c
                })
                data_rows += row_count
                data_rows_range += row_count
                data_columns += col_count

        if data_rows_range == 0:
            row = [b["value"] for b in data_blocks]
            if len(row) > 0:
                self.rows = [row]
            return

        if data_columns > 0:
            header_cols = len(self.final_header)
            if header_cols != data_columns:
                raise ValueError(f"Header / Data Column missmatch {header_cols} vs {data_columns}")

        rows = []
        values = self.sheet.values

        for row_offset in range(row_count):
            row = []
            for idx, entry in enumerate(data_blocks):
                if "static" in entry["type"]:
                    value = entry["value"]
                    row.append(value)
                elif "fixed" in entry["type"]:
                    value = entry["value"]
                    row.append(value)
                elif "range" in entry["type"]:
                    min_r = entry["min_r"]
                    min_c = entry["min_c"]
                    col_count = entry["col_count"]

                    for col in range(col_count):
                        src_c = min_c + col - 1
                        src_r = min_r + row_offset - 1
                        cell = values[src_r, src_c]
                        row.append(str(cell) if pd.notna(cell) else self.na)

            rows.append(row)
        self.rows = rows

    def _build_final_header(self):
        final_header = []

        for h_block in self.header_blocks:
            if h_block["type"] in ("static", "fixed"):
                final_header.append(h_block["value"])
            elif h_block["type"] == "range":
                max_w = h_block["col_count"]
                final_header.extend(h_block["values"][:max_w])
                if len(h_block["values"]) < max_w:
                    final_header.extend([""] * (max_w - len(h_block["values"])))
        self.final_header = final_header

    def extract(self):
        self._load_sheet()
        self._build_header_blocks()
        self._build_final_header()
        self._extract_data_rows()

        headers = self.final_header
        rows = self.rows

        return headers, rows


class ExcelExtractor:
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

    extractor = ExcelExtractor(config_file, output_dir)
    extractor.run()

if __name__ == "__main__":
    main()
