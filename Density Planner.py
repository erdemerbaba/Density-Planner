# -*- coding: utf-8 -*-
"""
Spyder Editor
"""

import numpy as np
import pandas as pd
print(np.__version__)
print(pd.__version__)



#1
#all chart
print("\n\n-----1-chart------")
chart=pd.read_excel("input.xlsx",skiprows=1)
print(chart)
chartnp=np.array(chart)
print("\n",chartnp)



#2
print("\n\n-----2-sorted chart------")
sorting=chart.sort_values(by="due date")
print(sorting)



#3
print("\n\n-----3-dublicate due date------")
seen={}
dupes=[]
counta=-1
countb=0
for x in sorting["due date"]:
    counta+=1
    if x not in seen:
        seen[x] = 1
    else:
        if seen[x] == 1:
            dupes.append(x)
            countb=counta
        seen[x] += 1
print(seen)
print(dupes)
print("dublicate column: ",countb)



#4
print("\n\n-----4-product------")
orders=sorting["orders"]
orders1=sorting["orders"]
orders2=sorting["orders"]
orders3=sorting["orders"]
orders1=np.array(orders1)
orders2=np.array(orders2)
orders3=np.array(orders3)
p1=sorting["p1"]
p2=sorting["p2"]
p3=sorting["p3"]
p11=np.array(p1)
p22=np.array(p2)
p33=np.array(p3)
print(orders)
print("old arrays")
print(p11)
print(p22)
print(p33)
while countb:
    if p11[countb] < p11[countb-1]:
        p11[countb],p11[countb-1]=p11[countb-1],p11[countb]
        orders1[countb],orders1[countb-1]=orders1[countb-1],orders1[countb]
    if p22[countb] < p22[countb-1]:
        p22[countb],p22[countb-1]=p22[countb-1],p22[countb]
        orders2[countb],orders2[countb-1]=orders2[countb-1],orders2[countb]
    if p33[countb] < p33[countb-1]:
        p33[countb],p33[countb-1]=p33[countb-1],p33[countb]
        orders3[countb],orders3[countb-1]=orders3[countb-1],orders3[countb]
    break
print("new arrays")
print(p11)
print(p22)
print(p33)
print(orders1)



#5
print("\n\n-----5-max capacity------")
a1=sorting["a1"][0]
a2=sorting["a2"][0]
a3=sorting["a3"][0]
print(f"a1: {a1}  , a2: {a2}  , a3:{a3}")



#6
print("\n\n-----6-order placement------")
total1=0
totala=0
total=orders1[0]+orders1[1]+orders1[2]+orders1[3]+orders1[4]
order=-1
if p11[0]<a1:
    total1=p11[0]
    totala=orders1[0]
    order=0
    if p11[0]+p11[1]<a1:
        total1=p11[0]+p11[1]
        totala=orders1[0]+orders1[1]
        order=1
        if p11[0]+p11[1]+p11[2]<a1:
            total1=p11[0]+p11[1]+p11[2]
            totala=orders1[0]+orders1[1]+orders1[2]
            order=2
            if p11[0]+p11[1]+p11[2]+p11[3]<a1:
                total1=p11[0]+p11[1]+p11[2]+p11[3]
                totala=orders1[0]+orders1[1]+orders1[2]+orders1[3]
                order=3
                if p11[0]+p11[1]+p11[2]+p11[3]+p11[4]<a1:
                    total1=p11[0]+p11[1]+p11[2]+p11[3]+p11[4]
                    totala=orders1[0]+orders1[1]+orders1[2]+orders1[3]+orders1[4]
                    order=4
print("completed product p1 in first day: ",total1)
print("finished order for p1 product in first day: ",totala)
while True:
    
    #p1 day 1
    if total1%a1 != 0:
        total = total[len(totala):]
        print("remaining order for p1 product: ",total)
        empty=a1-total1
        print("how many empty product left for first day: ",empty)
        total[1]
    half=p11[order+1]
    day=half-empty
    print("first day product:",total1+empty)
    print("second day remaining product",day)
    #p1 day 2
    day2=day-a1
    print("second day product:",day-day2)
    print("third day remaining product: ",day2)
    if len(total)<2 and day-day2<0:
        break
    totalday2=total
    
    #p1 day 3
    day3=day2-a1
    print("third day product:",day2-day3)
    print("fourth day remaining product: ",day3)
    if len(total)<2 and day2-day3<0:
        break
    totalday3=total
    
    #p1 day 4
    day4=day3-a1
    if day4<20 and len(total)>2:
        total = total[2:]
        res = None
        for i in range(0, len(orders1)): 
            if orders1[i] == total: 
                res = i + 1
                break
        if res == None:
            print ("finish")
            break
        else: 
            #print ("Character {} is present at {}".format(total, str(res))) 
            print("fourth day product:",(day3-day4)+p11[int(res)-1]-a1)
            print("fifth day remaining product: ",(day3+a1-p11[int(res)-1]))
    if len(total)<2 and day3-day4<0:
        break
    totalday4=total
    
    #p1 day 5
    day5=day4-a1
    if day5<20 and len(total)>2:
        total = total[2:]
        res = None
        for i in range(0, len(orders1)): 
            if orders1[i] == total: 
                res = i + 1
                break
        if res == None:
            print ("finish")
            break
        else: 
            #print ("Character {} is present at {}".format(total, str(res))) 
            print("fifth day product:",(day3+a1-p11[int(res)-1]))
            print("late-remaining product: ",((day4-day5)+p11[int(res)-1])-p11[int(res)-1])
    if len(total)<2 and day4-day5<0:
        break
    print("fifth day product:",(day3+a1-p11[int(res)-1]))
    if (day3+a1-p11[int(res)-1])>a1:
        print("late-remaining product: ",((day3+a1-p11[int(res)-1])-a1))
    totalday5=total
    break


#p2 day 1
#p2 day 2
#p2 day 3
#p2 day 4
#p2 day 5


#p3 day 1
#p3 day 2
#p3 day 3
#p3 day 4
#p3 day 5



#7
print("\n\n-----7-output template------\ndone")
import xlsxwriter 
workbook = xlsxwriter.Workbook('output.xlsx') 
worksheet = workbook.add_worksheet() 
worksheet.write('A1', 'DATES') 
worksheet.write('B1', 'p1') 
worksheet.write('C1', 'p2') 
worksheet.write('D1', 'p3') 
worksheet.write('A12', "LATE") 
worksheet.write('A13', 'NUMBER') 


#8&9
print("\n\n-----8- output order placement------\ndone")
outptdates=sorting["due date"]
outpta=np.array(outptdates)
#due dates
worksheet.write('A2',outpta[0] )
worksheet.write('A4',outpta[1] )
worksheet.write('A6',outpta[2] )
worksheet.write('A8',outpta[3] )
worksheet.write('A10',outpta[4] )

worksheet.write('B2',totala )
worksheet.write('B3',total1+empty )
worksheet.write('B4',totalday2[:2] )
worksheet.write('B5',day-day2 )
worksheet.write('B6',totalday3[:2] )
worksheet.write('B7',day2-day3 )
worksheet.write('B8',totalday3 )
worksheet.write('B9', (day3-day4)+p11[int(res)-1]-a1 )
worksheet.write('B10',totalday5[:2] )
worksheet.write('B11',(day3+a1-p11[int(res)-1] ))



#10
print("\n\n-----10- output ------\ndone")
#late number
if (day3+a1-p11[int(res)-1])-a1>0:
    worksheet.write('B13', a1-(day3+a1-p11[int(res)-1]))
    worksheet.write('B12', total )
else:
    worksheet.write('B13', "-")
    worksheet.write('B12', "-" )

workbook.close() 

#ER
