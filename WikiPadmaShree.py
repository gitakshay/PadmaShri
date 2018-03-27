from bs4 import BeautifulSoup
import urllib.request as req
import xlwt


def writetosheet(sheet,row,col,value):
    try:
        sheet.write(row,col,value)
    except Exception  as e:
        print("Unable to write data {} to sheet at {},{} encountered below error:\n".format(value,row,col)+str(e))
    return sheet
    
def getListToExcel(urls,workbook):
    sheet_list=workbook.add_sheet('List')
    xl_row=0
    for url in urls:
        print("URL ::"+url)
        page = req.urlopen(url)
        soup = BeautifulSoup(page,'html.parser')
        tables = soup.find_all('table')
        caption='List of Padma Shri award recipients, showing the year, field, and state/country[1]'
        for table in tables:
            #print(type(table.find('caption')))
            if not isinstance(table.find('caption'), type(None)):
                #print(table.find('caption').get_text())
                if table.find('caption').get_text()==caption:
                    padma_table=table
                    
        rows= padma_table.find_all('tr')
        for tr in rows[1:]:
            td=tr.find_all('td')
            name=tr.find('span',{"class":"fn"}).get_text()#.find_all('span')
            xl_row+=1
            year=td[0].get_text()
            field=td[1].get_text()
            state=td[2].get_text()
            print("Name:{} ,Year:{} ,Field:{} ,State:{} ".format(name,year,field,state))
            sheet_list=writetosheet(sheet_list,xl_row,0,name)
            sheet_list=writetosheet(sheet_list,xl_row,1,year)
            sheet_list=writetosheet(sheet_list,xl_row,2,field)
            sheet_list=writetosheet(sheet_list,xl_row,3,state)
        
    return workbook

url_list=['https://en.wikipedia.org/wiki/List_of_Padma_Shri_award_recipients_(1954%E2%80%931959)',
          'https://en.wikipedia.org/wiki/List_of_Padma_Shri_award_recipients_(1960%E2%80%931969)',
          'https://en.wikipedia.org/wiki/List_of_Padma_Shri_award_recipients_(1970%E2%80%931979)',
          'https://en.wikipedia.org/wiki/List_of_Padma_Shri_award_recipients_(1980%E2%80%931989)',
          'https://en.wikipedia.org/wiki/List_of_Padma_Shri_award_recipients_(1990%E2%80%931999)',
          'https://en.wikipedia.org/wiki/List_of_Padma_Shri_award_recipients_(2000%E2%80%932009)',
          'https://en.wikipedia.org/wiki/List_of_Padma_Shri_award_recipients_(2010%E2%80%932019)']

workbook=xlwt.Workbook()

workbook=getListToExcel(url_list,workbook)

workbook.save('Padmashree_List.xls')


