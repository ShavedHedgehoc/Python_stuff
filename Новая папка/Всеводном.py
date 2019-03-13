import os
import sys
import pandas as pd
import shutil

class MyErr(Exception): 
    def __init__(self,value):        
        self.value=value
    def __str__(self):
        return repr(self.value)
        

def get_files():
    work_directory=os.getcwd()
    files=os.listdir(work_directory)
    new_files=[]    
    for one_file in files:
        if os.path.splitext(one_file)[-1]=='.xlsx'\
           and os.path.splitext(one_file)[0].find('Взвешивание')!=-1:
            new_files.append(one_file)
    return(new_files)

def read_xl_file(r_file):
    xl = pd.ExcelFile(r_file)
    sh=xl.sheet_names[0]
    r_df=pd.read_excel(xl, sheet_name=sh, header=None, dtype=str)
    return r_df

def collect_files():
    files=get_files()
    all_append=False
    cur_directory=os.getcwd()
    ##dest_directory=cur_directory+'\\Архив\\'
    dest_directory=os.path.abspath("X:\Архив РП\АРХИВ ТЕКУЩИХ ОСТАТКОВ\Архив прослеживаемости\\")
    if not os.path.isdir(dest_directory):    os.makedirs(dest_directory)
    if os.path.exists(os.path.join(cur_directory, 'All.xlsx')):
        files.append('All.xlsx')
        all_append=True
    files_count=len(files)
    df_collect_files = None
    if files_count==0:
        pass        
    elif files_count==1 and all_append:
        df_collect_files=read_xl_file(files[0])
        if df_collect_files.shape[-1]<8: df_collect_files[7]=None        
        df_collect_files[7]=df_collect_files[7].fillna(0)
        df_collect_files[7] = df_collect_files[7].apply(pd.to_numeric)
        df_collect_files.drop_duplicates(inplace=True)
    else:
        count_passed=0
        count_removed=0
        for ff in files:
            df_temp=read_xl_file(ff)
            if df_temp.iloc[0,1]!='{Document.ШтрихкодЕмкости}':
                df_collect_files=pd.concat([df_collect_files,df_temp],ignore_index=True)
                shutil.move(os.path.join(cur_directory, ff),os.path.join(dest_directory, ff))
            else:
                os.remove(os.path.join(cur_directory, ff))
                count_removed+=1
            count_passed+=1
            os.system("cls")
            print("Обработано файлов %d из %d. Пропущено файлов %d" %(count_passed,len(files),count_removed))
            sys.stdout.flush()
            os.system("cls")
        if df_collect_files.shape[-1]<8: df_collect_files[7]=None        
        df_collect_files[7]=df_collect_files[7].fillna(0)
        df_collect_files[7] = df_collect_files[7].apply(pd.to_numeric)
        df_collect_files.drop_duplicates(inplace=True)    
        print("Записываю файл архива...")
        df_collect_files.to_excel('All.xlsx',index=False,header=False)    
    return df_collect_files

def group_by_batch(df):
    if df is None or df.empty:
        df_data=None
    else:
        df_first_group=df[[2,3,4,5,6,7]].groupby([2,3])
        df_data=pd.DataFrame(dtype=str)
        for name, group in df_first_group:
            t_list=[]
            df_second_group=group.groupby([5])[7].sum()
            t_list.append(name[0])
            t_list.append(name[1])        
            for i in df_second_group.index:        
                t_list.append(i)
                t_list.append(df_second_group[i])
            t_series=pd.Series(t_list)    
            df_data=df_data.append(t_series,ignore_index=True)
    return df_data

if __name__ == "__main__":
    
    try:
        if not os.path.exists(os.path.join(os.getcwd(), 'Варки.xlsx')):raise MyErr(0)
        files=get_files()
        df_collect=collect_files()    
        if df_collect is None: raise MyErr(1)    
        if df_collect.empty: raise MyErr(2)
        print("Обрабатываю данные...")
        df_res=group_by_batch(df_collect)
        df_base=read_xl_file('Варки.xlsx')
        if df_base.empty: raise MyErr(3)
        old_df_size=df_base.shape[1]
        new_df_size=df_base.shape[1]+df_res.shape[1]-1
        df_base=df_base.reindex(columns=range(new_df_size))
        for row in df_res.itertuples():
            row_find=df_base.index[(df_base[6].map(lambda x: '-' if x=='-' else int(x))==int(row[2])) & (df_base[3]== row[1])].tolist()
            if len(row_find)!=0:
                for app_row in row_find:
                    ttt=pd.Series(row).tolist()
                    df_base.loc[app_row,range(old_df_size+1,new_df_size)]=ttt[3:]
        
        df_base = df_base.replace('nan', '', regex=True)            
        df_base.to_excel('Result.xlsx',index=False,header=False)
    
    except MyErr as e:
        if e.value==0:
            print('В папке нет файла варок...')
            print('Поместите файл с именем "Варки.xlsx" в папку и запустите программу еще раз.')
        if e.value==1:
            print('В папке нет файлов сканирования или файла архива...')
            print('Поместите файлы в папку и запустите программу еще раз.')
        if e.value==2:
            print('Файл архива пустой...')
            print('Удалите файл "All.xlsx" и запустите программу еще раз.')
        if e.value==3:
            print('Файл варок пустой...')
            print('Поместите аполненный файл с именем "Варки.xlsx" в папку и запустите программу еще раз.')
    finally:
        print('Рассчет окончен')
