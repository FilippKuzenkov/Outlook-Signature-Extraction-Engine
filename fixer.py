# Re-interpret the broken text as Latin-1, convert back to UTF-8
import pandas as pd

with open("your_csv", "r", encoding="latin1") as f:
    df = pd.read_csv(f, sep=";")  # adjust separator if needed

df.to_csv("your_csv", sep=";", index=False, encoding="utf-8")