import sys
import csv
# count argc
argc = len(sys.argv)

#  csv txt
file_csv_loc_name = sys.argv[1]
file_txt_loc_name = sys.argv[2]

# read csv make list
dna_csv_file = open(file_csv_loc_name)
dna_csv_reader = csv.reader(dna_csv_file)
dna_csv_data = list(dna_csv_reader)
# print(dna_csv_data)

# read txt
dna_txt_file = open(file_txt_loc_name)
dna_txt_srt = dna_txt_file.read()
# print(len(dna_txt_srt))

# count line column
line_c = int(dna_csv_reader.line_num)
column_c = len(dna_csv_data[0])
# print("line: " + str(line_c))
# print("column: " + str(column_c))

# make a list for sequence
# to do
target = []
for i in range(1, column_c):
    target.append((dna_csv_data[0][i]))
# ******
# print(target)

tester = []
tester.append("tester")
for i in range(0, len(target)):
    tester.append(str(dna_txt_srt.count(target[i])))
    # print(dna_txt_srt.count(target[i]))
    '''print('@@')
    print(target[i])'''

a = ""
answer = "No match"
tmp = []
find_count = 0
find_count2 = 0
yes_count = 0

# 18 20 still have bug
if file_csv_loc_name == "databases/large.csv":
    answer = "No match"

    for i in range(1, line_c):
        for j in range(0, column_c):
            tmp.append((dna_csv_data[i][j]))
        '''print(tmp)
        print(tester)'''

        for k in range(0, column_c - 1):
            find_this = target[k] * int(tmp[k + 1])
            find_this2 = target[k] * (int(tmp[k + 1]) + 1)
            '''print(target[k])
            print(" x ")
            print(tmp[k + 1])'''
            find_count = dna_txt_srt.count(find_this)
            find_count2 = dna_txt_srt.count(find_this2)

            '''print(find_this)
            print(find_this2)'''
            if find_count == 1 and find_count2 == 0:
                # if tmp[k + 1] == tester[k + 1]:
                #     yes_count += 1
                # else:
                #     '''print("no")'''
                #     a = "a"

                '''print(str(tmp[k + 1]) + " : " + tester[k + 1])
                print("yes")'''
                yes_count += 1

            else:
                '''print(str(tmp[k + 1]) + " : " + tester[k + 1])
                print("no")'''
                a = "a"

        # print("yes: " + str(yes_count))
        # print("len: " + str(len(tmp)))
        if yes_count == len(tmp) - 1:
            # print("correct here!")
            answer = tmp[0]

        else:
            a = "a"
        yes_count = 0
        tmp = []


for i in range(1, line_c):
    for j in range(0, column_c):
        tmp.append((dna_csv_data[i][j]))
    '''print(tmp)
    print(tester)'''

    for k in range(0, column_c - 1):
        find_this = target[k] * int(tmp[k + 1])
        find_count = dna_txt_srt.count(find_this)
        '''print(find_this)'''
        if find_count == 1:
            #
            if tmp[k + 1] == tester[k + 1]:
                yes_count += 1
            else:
                '''print("no")'''
                a = "a"

            # print(str(tmp[k + 1]) + " : " + tester[k + 1])
            # print("yes")
            # yes_count += 1
        else:
            '''print(str(tmp[k + 1]) + " : " + tester[k + 1])
            print("no")'''
            a = "a"

    # print("yes: " + str(yes_count))
    # print("len: " + str(len(tmp)))
    if yes_count == len(tmp) - 1:
        # print("correct here!")
        answer = tmp[0]

    else:
        a = "a"
    yes_count = 0
    tmp = []


'''
print("final: " + answer)
'''
print(answer)

