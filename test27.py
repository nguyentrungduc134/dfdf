import openpyxl
wb = openpyxl.load_workbook('F:\\automation\\0627rp.xlsx')
sheet = wb.active
wb2 = openpyxl.load_workbook('F:\\automation\Daily_report.xlsx')
sheet2 = wb2.active

for y in range(12, 61):
     cell1 = "E%d" % y
     y1 = sheet2[cell1].value
     print(y1)
     x = 2
     z = 9
     while True:
        cell = "H%d" % x
        x1 = sheet[cell]
        if x1.value:        
            print(x1.value)
            if x1.value == y1:
                 cell5 = "N%d" % z
                 a1 = sheet2[cell5]
                 a1.value = y1
                 cell1 = "M%d" % z
                 a2 = sheet2[cell1]
                 c1 = "K%d" % x
                 a2.value = sheet[c1].value
                 cell2 = "L%d" % z
                 a3 = sheet2[cell2]
                 c2 = "E%d" % x
                 a3.value = sheet[c2].value
                 cell3 = "O%d" % z
                 a4 = sheet2[cell3]
                 c3 = "R%d" % x
                 a4.value = sheet[c3].value

                 cell4 = "P%d" % z
                 a5 = sheet2[cell4]
                 c4 = "Q%d" % x
                 text = "Object:%s" % sheet[c4].value
                 a5.value = text
                 
                 cell4 = "P%d" % z
                 a5 = sheet2[cell4]
                 c4 = "O%d" % x
                 text = "Class:%s" % sheet[c4].value
                 new = a5.value+"\n%s" % text
                 a5.value = new
                 
                 cell4 = "P%d" % z
                 a5 = sheet2[cell4]
                 c4 = "AI%d" % x
                 text = "Message:%s" % sheet[c4].value
                 new = a5.value+"\n%s" % text
                 a5.value = new

                 cell4 = "P%d" % z
                 a5 = sheet2[cell4]
                 c4 = "P%d" % x
                 text = "Message:%s" % sheet[c4].value
                 new = a5.value+"\n%s" % text
                 a5.value = new


                 wb2.save('F:\\automation\Daily_report.xlsx') 
                 print(y)
        else:
             print("not ok")
             break
        x = x + 1
        z = z + 1 
