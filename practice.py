import pandas as pd

#엑셀 읽기
a1 = pd.read_excel('(Yoon) 완료 68회 PreStarter 속성분석의 사본.xlsx', header=1)
#읽은 엑셀을 리스트로변환
alist = a1.values.tolist()

#리스트에서 한 행씩 읽어서 str변수에 원하는 형태로 삽입
tupletext = "sign = {"
for i in range(len(alist)):
    tupletext += "\""+str(alist[i][2])+"\" : " + str(alist[i][1]) + ","
    if i%7 == 0 :tupletext += "\n"

tupletext += "}"

#text 파일 만들기
f= open("sign2.txt","w",-1,encoding="utf-8")
#만든 str을 text에 넣기
f.write(tupletext)
#파일 닫기
f.close()
print(tupletext)