import pandas as pd
from dotenv import load_dotenv
import os

load_dotenv()

def add_df_to_excel(fpath, new_df, sheet_name='Sheet1'):
    if os.path.exists(fpath):
        try:
            existing_df = pd.read_excel(fpath, sheet_name=sheet_name)
            start_row = existing_df.shape[0] + 1  
            writer_mode = 'a'
            sheet_exists = 'overlay'
            header = False 
            print(f"파일 '{fpath}'이(가) 존재합니다. 데이터가 {start_row}행부터 추가됩니다.")
        except ValueError:
            start_row = 0
            writer_mode = 'a'
            sheet_exists = 'overlay' 
            header = True 
            print(f"시트 '{sheet_name}'이(가) 존재하지 않습니다. 시트를 새로 생성하고 헤더를 포함합니다.")
    else:
        start_row = 0
        writer_mode = 'w' 
        sheet_exists = None 
        header = True 
        print(f"파일 '{fpath}'을(를) 찾을 수 없습니다. 새로운 파일로 생성하고 헤더를 포함합니다.")

    with pd.ExcelWriter(
        fpath, 
        engine='openpyxl', 
        mode=writer_mode,
        if_sheet_exists=sheet_exists 
    ) as writer:
        new_df.to_excel(
            writer, 
            sheet_name=sheet_name, 
            startrow=start_row, 
            index=False,        
            header=header 
        )

if __name__ == '__main__':
    data_dir = os.getenv('data_dir')
    fille_name = 'existing_file.xlsx' 
    file = os.path.join(data_dir, fille_name)

    data = {'Name': ['kohenil', 'ihje'], 'Age': [58, 33]}
    df = pd.DataFrame(data)
    add_df_to_excel(file,df)