import pandas as pd

if __name__ == '__main__':
    df = pd.read_excel("C:\\Users\\PC\\Downloads\\매입전자세금계산서목록(1~26).xls", header=5)
    base_df = pd.read_excel("C:\\Users\\PC\\Documents\\python\\account\\semoosa\\data\\25장부.xlsx", sheet_name='25년장부', header=3)

    #start_date = df['작성일자'].min()
    #end_date = df['작성일자'].max()
    print(base_df.columns)

    base_df['날짜'] = pd.to_datetime(base_df['날짜'])

    #filtered_bs_df = base_df[(base_df['날짜'] >= start_date) & (base_df['날짜'] <= end_date) & (base_df['계정과목'].str.contains('부가세'))]
    filtered_bs_df = base_df[(base_df['날짜'].dt.quarter == 3 ) & (base_df['계정과목'].str.contains('부가세'))]

    filtered_bs_df.to_csv('abc.csv', index=False, encoding='utf-8-sig')

    print(df.columns)
    print(filtered_bs_df[['날짜','계정과목']])
