# -*- coding: utf-8 -*-
"""
Created on Sun Apr 12 12:10:59 2020

@author: trant
"""


from openpyxl import load_workbook
import matplotlib.pyplot as plt
import numpy as np

"""---------------------- ADT: SINGLY LINKED LIST ----------------------"""
class Node:
    def __init__(self, tinh, data):
        self.tinh = tinh
        self.data = data
        self.next = None

class SList:
    def __init__(self, node = None):
        self.head = node
        self.tail = node
        self.size = 0

    #   In toàn bộ danh sách. \\ không in Node đầu
    def printTinh1(self):
        cur = self.head.next
        print('+ {:-^6} + {:-^15} + {:-^16} +'.format('', '', ''))
        print('| {:^6} | {:^15} |{:^16}|'.format('STT', 'Tên Tỉnh','Thu Ngân Sách(TrĐ)'))
        print('+ {:-^6} + {:-^15} + {:-^16} +'.format('', '', ''))
        stt = 0
        while cur != None:
            stt += 1
            print('| {:^6} | {:^15} | {:^16} |'.format(stt, cur.tinh, cur.data))
            cur = cur.next
        print('+ {:-^6} + {:-^15} + {:-^16} +'.format('', '', ''))

    #   In danh sách có điều kiện.
    def printTinh2(self, k = None):
        cur = self.head
        print('+ {:-^6} + {:-^15} + {:-^10} +'.format('', '', ''))
        print('| {:^6} | {:^15} |{:^10}  |'.format('STT', 'Tên Tỉnh', ''))
        print('+ {:-^6} + {:-^15} + {:-^10} +'.format('', '', ''))
        stt = 0
        while stt != k:
            stt += 1
            print('| {:^6} | {:^15} | {:>10} |'.format(stt, cur.tinh, cur.data))
            cur = cur.next
        print('+ {:-^6} + {:-^15} + {:-^10} +'.format('', '', ''))

    #   Hàm tính tổng data.
    def sum(self):
        total = 0
        cur = self.head
        while cur != None:
            total += cur.data
            cur = cur.next
        return total

    # Hàm thêm Node vào đầu danh sách -- O(1)
    def appendHead(self, newNode):
        if self.head is None:
            self.head = newNode
            self.tail = newNode
            self.size = 1
        newNode.next = self.head
        self.head = newNode
        self.size += 1

    #   Hàm thêm Node vào cuối danh sách -- O(1)
    def appendTail(self, newNode):
        if self.head is None:
            self.head = newNode
            self.tail = newNode
            self.size = 1
        self.tail.next = newNode
        self.tail = newNode
        self.size += 1

    #   Hàm chèn 1 Node mới vào sau Node trong danh sách -- O(1)
    def insertAfter(self, Node, newNode):
        self.size += 1
        newNode.next = Node.next
        Node.next = newNode

    #   Hàm thêm 1 Node đồng thời sắp xếp danh sách -- O(N)
    def AppendAndSort(self, tinh, data):
        newNode = Node(tinh, data)
        cur = self.head
        if self.head is None:
            self.head = newNode
            self.tail = newNode
        elif newNode.data > self.head.data:
            self.appendHead(newNode)
        else:
            if newNode.data <= self.tail.data:
                self.appendTail(newNode)
            if newNode.data < self.head.data and newNode.data > self.tail.data:
                while cur is not self.tail:
                    if newNode.data < cur.data and newNode.data > cur.next.data:
                        self.insertAfter(cur, newNode)
                    cur = cur.next

    def get_value_excel(self, rowStart, rowEnd, col1 = 'C', col2 = 'ZZ', bool = True):
        wb = load_workbook("So lieu du toan NSNN nam 2019.xlsx")
        sheet = wb["B21"]
        for i_row in range(rowStart, rowEnd):
            cell_tinh = "%s%s" % ('B', i_row)
            name_tinh = sheet[cell_tinh].value

            cell_col1 = "%s%s" % (col1, i_row)
            cell_col2 = "%s%s" % (col2, i_row)
            val1 = sheet[cell_col1].value
            val2 = sheet[cell_col2].value
            if val2 is None:
                val2 = 0
            if bool == True:
                val = val1 + val2
            else:
                val = val1 - val2
            self.AppendAndSort(name_tinh, val)
        wb.close()

#   -------------------------------------------------------------

# -- Dữ liệu của 6 vùng và các tỉnh thuộc 6 vùng.
vung1 = SList()
vung1.get_value_excel(14, 29, 'C', 'Z', True)
vung2 = SList()
vung2.get_value_excel(29, 41, 'C', 'Z', True)
vung3 = SList()
vung3.get_value_excel(41, 56, 'C', 'Z', True)
vung4 = SList()
vung4.get_value_excel(56, 62, 'C', 'Z', True)
vung5 = SList()
vung5.get_value_excel(62, 69, 'C', 'Z', True)
vung6 = SList()
vung6.get_value_excel(69, 83, 'C', 'Z', True)

# -- Dữ liệu tổng thu ngân sách của 63 tỉnh/TP.
data1 = SList()
data1.get_value_excel(15, 29, 'C', 'Z', True)
data1.get_value_excel(30, 41, 'C', 'Z', True)
data1.get_value_excel(42, 56, 'C', 'Z', True)
data1.get_value_excel(57, 62, 'C', 'Z', True)
data1.get_value_excel(63, 69, 'C', 'Z', True)
data1.get_value_excel(70, 83, 'C', 'Z', True)

# -- Dữ liệu nhà nước bổ xung ngân sách cho địa phương.
data2 = SList()
data2.get_value_excel(15, 29, 'I', 'J', True)
data2.get_value_excel(30, 41, 'I', 'J', True)
data2.get_value_excel(42, 56, 'I', 'J', True)
data2.get_value_excel(57, 62, 'I', 'J', True)
data2.get_value_excel(63, 69, 'I', 'J', True)
data2.get_value_excel(70, 83, 'I', 'J', True)

# -- Dữ liệu các tỉnh 'đóng góp' vào ngân sách nhà nước.(column C - D)
data3 = SList()
data3.get_value_excel(15, 29, 'C', 'D', False)
data3.get_value_excel(30, 41, 'C', 'D', False)
data3.get_value_excel(42, 56, 'C', 'D', False)
data3.get_value_excel(57, 62, 'C', 'D', False)
data3.get_value_excel(63, 69, 'C', 'D', False)
data3.get_value_excel(70, 83, 'C', 'D', False)
#   -------------------------------------------------------------

#   Vùng code xử lý vẽ biểu đồ.
fig, ax = plt.subplots(figsize=(10, 5), subplot_kw=dict(aspect="equal"))

name_vung = [vung1.head.tinh, vung2.head.tinh, vung3.head.tinh,
            vung4.head.tinh, vung5.head.tinh, vung6.head.tinh]
nsnn = [vung1.head.data, vung2.head.data, vung3.head.data,
        vung4.head.data, vung5.head.data, vung6.head.data]
wedges, texts, autotexts = ax.pie(nsnn ,autopct='%1.2f%%', shadow=True)
ax.legend(wedges, name_vung,
          title="CÁC VÙNG",
          loc="center left",
          bbox_to_anchor=(1, 0, 0.5, 1))
plt.setp(autotexts, size=7, weight="bold")
ax.set_title(" BIỂU ĐỒ TRÒN THỂ HIỆN THU NGÂN SÁCH CỦA CÁC VÙNG")


print("1. Vẽ biểu đồ tròn biểu diễn tỉ lệ thu ngân sách của các vùng.")
plt.show()
    

print("\n2. Nhập vào số k, trả về k tỉnh có thu ngân sách cao nhất theo thứ tự giảm dần.")  
k = int(input('k = '))
data1.printTinh2(k)
    

print("\n3. Nhập vào số k, trả về k tỉnh nhận được bổ sung từ ngân sách nhà nước nhiều nhất theo thứ tự giảm dần.") 
k = int(input('k = '))
data2.printTinh2(k)
    
    
print("\n4. Nhập vào số k, trả về k tỉnh có đóng góp vào ngân sách nhà nước cao nhất.")
k = int(input('k = '))
data3.printTinh2(k)


print("\n5. Nhập vào tên một vùng, trả về tên các tỉnh trong vùng cùng tổng thu ngân sách của tỉnh đó theo thứ tự giảm dần. ")
tenvung = input('Nhập tên vùng(in hoa, có dấu): ')
if tenvung == vung1.head.tinh:
    vung1.printTinh1()
elif tenvung == vung2.head.tinh:
    vung2.printTinh1()
elif tenvung == vung3.head.tinh:
    vung3.printTinh1()
elif tenvung == vung4.head.tinh:
    vung4.printTinh1()
elif tenvung == vung5.head.tinh:
    vung5.printTinh1()
elif tenvung == vung6.head.tinh:
    vung6.printTinh1()
else:
    print("Tên vùng không hợp lệ!")


print("\n6. Liệt kê tên các tỉnh (ít nhất) có tổng nguồn thu ngân sách chiếm tới 50% tổng thu của cả nước.")
k = 0.5 * sum(nsnn)
total = 0
x = Node(0, 0)
newlist = SList(x)

while data1.head.data + total < k:
    cur = data1.head
    data1.head = cur.next
    cur.next = None
    total += cur.data
    newlist.appendTail(cur)

cur = data1.head
while cur != None:
    if cur.data < k - total:
        val = cur
        cur = val.next
        val.next = None
        newlist.appendTail(val)
        total += val.data
    if k - total < 0:
        break
    cur = cur.next

newlist.printTinh1()