#range 사용하기 
#여러 개의 셀이 포함되어 있는 range를 사용하기
import openpyxl

wb = openpyxl.Workbook()
sheet = wb.active
a = sheet['A1':'B9']# 범위를 불러왔을 때, 튜플이 이중으로 되어있음.하나의 행에 대한 정보가 튜플로 만들어져있음, 행을 모아놓은 것을 또 튜플로 가지고 있음
# ((<Cell 'Sheet'.A1>, <Cell 'Sheet'.B1>, <Cell 'Sheet'.C1>), (<Cell 'Sheet'.A2>, <Cell 'Sheet'.B2>, <Cell 'Sheet'.C2>), (<Cell 'Sheet'.A3>, <Cell 'Sheet'.B3>, <Cell 'Sheet'.C3>))

no = 1
for x,y in a:
    print(x, y)
    x.value = 2
    y.value = no
    no += 1

for x,y in a:
    print(x.value, y.value)