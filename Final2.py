def verify_url(url):
    import requests
    from requests.exceptions import MissingSchema
    try:
        request = requests.get(url)
    except MissingSchema:
        return 0;
    return 1;

def data_scrap(livedata):
    import urllib3
    from bs4 import BeautifulSoup
    http = urllib3.PoolManager()
    response2 = http.request('GET', livedata)
    soup2 = BeautifulSoup(response2.data.decode('utf-8'), features='html.parser')
    return soup2

def data_cleaning(soup2):
    import numpy as np
    import re
    np_str=np.str(soup2)
    match = re.search(r'(\d+/\d+/\d+)',np_str)
    if match==None:
#         print('No data found.')
        np_str3='0'
    else:
        np_str2=np_str[match.span()[0]:].replace('\t',' ')
        while re.search('  ',np_str2)!=None:
            np_str2=np_str2.replace('  ',' ')
        np_str3=np_str2.replace(' ',';').split('<br/>')
    return np_str3

def dataframe(np_str3):
    import pandas as pd
    df=pd.DataFrame({"Data": np_str3})['Data'].str.split(';',expand=True)
    for i in range(0,len(df.columns)):
        if sum(df[i]=='')>=len(df[i])/1.01:
            df.drop(i,axis=1,inplace=True)
    df_cols=["Date","Time","Price","Size","Indicator"]
#     print(df)
    if len(df_cols)==len(df.columns):
        df.columns=df_cols
    return df

def find_xlsx(dir_path, suffix):
    import os
    print(suffix)
    filenames = os.listdir(dir_path)
    files=[filename for filename in filenames if ((filename.endswith(suffix[0]) or filename.endswith(suffix[1]) or filename.endswith(suffix[2])) and not(filename.startswith("~")))]
    return files

def save_to_excel(writer,df_final,Sheet_name):
    import numpy as np
    df_final.to_excel(writer,sheet_name=Sheet_name,header=False,index=False)
    return 0
def main(livedata):
    soup2=data_scrap(livedata)
    np_str3=data_cleaning(soup2)
    if np_str3=='0':
        df='0'
    else:
        df=dataframe(np_str3)
    return df
def filename(i,Folder_name):
    import datetime
    now = datetime.datetime.now().strftime("%Y%m%d")
    fn = i.replace('xlsm','xlsx')
    if os.path.isfile("./"+Folder_name+"/"+now+"_"+fn):
        now = datetime.datetime.now().strftime("%Y%m%d%H%M")
        fn2 = now+"_"+fn
    else:
        fn2 = now+"_"+fn
    return fn2
def foldername(i):
    import os
    Folder_name=i.replace('.xlsm','')
    if os.path.isdir("./"+Folder_name)==False:
        os.mkdir("./"+Folder_name)
    return Folder_name
import warnings
warnings.filterwarnings("ignore")
import pandas as pd
import numpy as np
import os
os.chdir(os.getcwd())
dir_path=os.getcwd()
suffix=[".xlsx",".xls",".xlsm"]
files=find_xlsx(dir_path,suffix)
vali=0;
for i in files:
    df=pd.read_excel(i) # Read Excel 
    Folder_name=foldername(i)
    File_name=filename(i,Folder_name)
    Sheet_names=df.Names
    val=0
    vali=vali+1
    writer = pd.ExcelWriter(Folder_name+"/"+File_name, engine='xlsxwriter')
    df.to_excel(writer,sheet_name="Links")
    print("Working on File: "+i+"\n")
    if "Links" in df: # Check if links column exists in the dataframe
        for livedata in df.Links:
            if verify_url(livedata)==0:
                print('Invalid URL: ',livedata)
            else:
                df_final=main(livedata)
                if type(df_final)==str:
                    print(Sheet_names[val]+" sheet is created as empty.")
                else:
                    save_to_excel(writer,df_final,Sheet_names[val])
                    print(Sheet_names[val]+" sheet is created.")
            val=val+1
    writer.save()
exit(0)