import pandas as pd
path = r'D:\project\HR데이터\데이터\역량진단 데이터\역량진단_응답데이터\23년_하반기_역량진단_응답데이터.xlsx'
df = pd.read_excel(path, sheet_name='문항 정보_타인진단')
print(df.columns.tolist())
print(df.head().to_string())
