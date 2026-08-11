import pandas as pd
from pathlib import Path


BASE_DIR = Path('../' + input('>>> ')).resolve()
OUTPUT_DIR = BASE_DIR / "COMPLETED"
OUTPUT_DIR.mkdir(exist_ok=True)



print(f'{BASE_DIR = }')
print(f'{OUTPUT_DIR = }')