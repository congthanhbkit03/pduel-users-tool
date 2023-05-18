import openpyxl
# file csv upload users
# username, firstname, lastname, password, email

#loafd file excel
wb = openpyxl.load_workbook('macle.xlsx')
sheet = wb.active   #lay sheet duoc active

danhsachAccount = []
#duyet qua cac hang cua file excel de doc - tu hang muon doc den het du lieu trong sheet
for row in range(6, sheet.max_row + 1):
    accObj = {}
    accObj['username'] = (sheet['C' + str(row)].value).lower()
    accObj['firstname'] = "{}_{}".format(sheet['AA' + str(row)].value, sheet['D' + str(row)].value)
    accObj['lastname'] = (sheet['E' + str(row)].value)
    accObj['email'] = accObj['firstname'] + "@pdu.edu.vn"
    accObj['password'] = accObj['username']

    #them vao danh sach
    danhsachAccount.append(accObj)

print(len(danhsachAccount))

# ghi danhsach ra file csv - upload
with open('dct22.csv', 'w', encoding="utf-8") as fw:
    # ghi dong dau tien vao file
    fw.write('username,firstname,lastname,password,email\n')
    # doc du lieu trong danhsachAccount va in ra moi account 1 dong
    for acc in danhsachAccount:
        line = "{},{},{},{},{}\n".format(acc['username'], acc['firstname'], acc['lastname'], acc['password'], acc['email'])
        # ghi line nayf vao file
        fw.write(line)


# ghi ra file csv - enrolment
with open('enrol_dct22.csv', 'w', encoding="utf-8") as fw:
    fw.write('account\n')
    for acc in danhsachAccount:
        fw.write(acc['username']+"\n")
