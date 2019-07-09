import pandas as pd
from docx.api import Document

from tqdm import tqdm

document = Document('plan_obs/bsa_dkr_2019_3.docx')
table = document.tables[0]

data = []

keys = None
for i, row in tqdm(enumerate(table.rows)):
    text = (cell.text for cell in row.cells)

    if i == 0:
        keys = tuple(text)
        continue
    row_data = dict(zip(keys, text))
    data.append(row_data)

df = pd.DataFrame(data)
df.to_csv('test.csv', sep=" ", header=True, index=False)
