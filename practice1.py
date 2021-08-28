dic = {}
list_str = ["a","b"]
list_num= [1, 2, 3, 4, 5, 6, 7, 8, 9]
list_tennum = [11,22,33,44,55,66,77,88,99]

list2 = [list_num, list_tennum]

for i in range(2):
    dic[list_str[i]] = list2[i]
    

print(dic)
