
def jumpNum():
    a = []
    i = 1
    j = 10
    while i < 10/2:
       a.append(i)
       a.append(j-i)
       i = i + 1
    a.append(i)
    return a

b = jumpNum()

for i in range(2, 10):
    for j in range(0, 9):
        print(str(i) + ' x ' + str(b[j]) + ' = ' + str(i * b[j]))
    print('----------')

input()
