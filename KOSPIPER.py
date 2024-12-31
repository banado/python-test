{
 "cells": [
  {
   "cell_type": "code",
   "execution_count": 21,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "Every 1 week at 17:58:00 do job() (last run: [never], next run: 2024-12-31 17:58:00)"
      ]
     },
     "execution_count": 21,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "from urllib.request import urlopen\n",
    "from urllib.parse import urlencode, unquote, quote_plus\n",
    "import urllib\n",
    "import requests\n",
    "import pandas as pd\n",
    "import json\n",
    "from datetime import datetime as dt\n",
    "import requests\n",
    "from datetime import datetime\n",
    "from dateutil.relativedelta import relativedelta\n",
    "from gspread_dataframe import get_as_dataframe, set_with_dataframe\n",
    "import gspread\n",
    "import schedule\n",
    "import time\n",
    "import webbrowser\n",
    "\n",
    "\n",
    "url='https://ecos.bok.or.kr/api/StatisticTableList/Q3DXRSAGZ4XDQH6SQEPU/json/kr/1/999'\n",
    "\n",
    "res=requests.get(url)\n",
    "dflist=pd.DataFrame(res.json()['StatisticTableList']['row'])\n",
    "code=dflist[dflist['STAT_NAME']=='1.5.1.2. 주식시장(월,년)']['STAT_CODE'].values[0]\n",
    "\n",
    "today = datetime.today()\n",
    "mo=today.month\n",
    "year=today.year\n",
    "\n",
    "\n",
    "url3=f'https://ecos.bok.or.kr/api/StatisticSearch/Q3DXRSAGZ4XDQH6SQEPU/json/kr/1/999/{code}/M/200001/{year}{mo}/1110000'\n",
    "res3=requests.get(url3)\n",
    "df=pd.DataFrame(res3.json()['StatisticSearch']['row'])\n",
    "df=df[['ITEM_NAME1','TIME','DATA_VALUE']]\n",
    "name=df['ITEM_NAME1'].values[0]\n",
    "df=df[['TIME','DATA_VALUE']]\n",
    "df.columns=['TIME','KOSPI PER']\n",
    "df['TIME']=df['TIME'].apply(lambda x:str(x)[:4]+'년'+' '+str(x)[4:]+'월')\n",
    "df['Average']=df['KOSPI PER'].astype(float).mean()\n",
    "\n",
    "\n",
    "json_key =\"spreadsheet-443708-db4b357189f2.json\"\n",
    "gc = gspread.service_account(json_key)\n",
    "spreadsheet_url = \"https://docs.google.com/spreadsheets/d/1x392Sxm1pRCCjwUx_sGNJB81z41FhipdbLDpSLCcTg8/edit?gid=221383335#gid=221383335\"\n",
    "doc = gc.open_by_url(spreadsheet_url)\n",
    "\n",
    "try:\n",
    "    worksheet = doc.worksheet('코스피PER')\n",
    "except Exception as sheet_name:\n",
    "    raise Exception(f'{sheet_name} 워크시트를 찾을 수 없음')\n",
    "    \n",
    "worksheet.clear()\n",
    "set_with_dataframe(worksheet=worksheet, dataframe=df, include_index=False,\n",
    "                              include_column_header=True, resize=True)\n",
    "\n"
   ]
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python [conda env:base] *",
   "language": "python",
   "name": "conda-base-py"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.12.7"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 4
}
