import openpyxl

dest_filename = './myworkbook.xlsx' 
wb = openpyxl.Workbook() # Workbook은 클래스고 적어도 한 개 이상의 worksheet를 만들어낸다. active를 통해 활성화시킬 수 있다.
sheet = wb.active #저장시킨 워크북을 wb.active를 통해서 시트를 활성화시킴. 
sheet.title = 'mysheet1' #시트의 이름을 설정함.
sheet['A1'].value = 'name' #시트에 columns가 A, rows가 1인 셀을 만들어냄.
sheet['A2'].value = 'yoon so young'
sheet.cell(row=3, column=1).value = 20
wb.save(dest_filename) #엑셀을이름을 저장해주는 것.저장해 주는 것. 

