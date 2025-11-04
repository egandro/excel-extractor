import os, argparse, pandas as pd

parser = argparse.ArgumentParser()
parser.add_argument("-o", "--output", default=".")
args = parser.parse_args()

os.makedirs(args.output, exist_ok=True)
out_file = os.path.join(args.output, "pattern.xlsx")

rows, cols = 20, 14
df = pd.DataFrame([[f"col{c}-row{r}" for c in range(cols)] for r in range(rows)],
                  columns=[f"header{i}" for i in range(cols)])
df.to_excel(out_file, index=False)