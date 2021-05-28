import re
import xlwt 
from xlwt import Workbook
from tkinter import filedialog
from tkinter import *

root = Tk()
root.filename =  filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("txt files","*.txt"),("all files","*.*")))

wb = Workbook() 


lines = open(root.filename).readlines()

basliklar = []
satirlar = []
degerler = []
indis = -1
mod = 0

for line in lines:
	indis += 1
	line = line.replace(',','.')
	#if re.match(r'^-?\d+(?:\.\d+)$', line) is None: # is float control 
	#if indis<4:	
	#	basliklar.append([line])
	#	continue
	#if indis==5:
	#	continue
	#else:
	#	if not line.strip():
	#		continue
	#	elif not line.strip():
	#		continue
	#	else:
	if not line.strip(): 		#boş satırların indisini satirlar dizisine atıyor
		satirlar.append(indis)
		continue	
	else:				#boş olmayan tüm satırları degerler dizisine atıyor
		sayilar = line.split()
		degerler.append(sayilar)
		
		

degerler_sayisi = len(degerler)	#degerler dizisinin eleman sayısını buluyor (yalnızca ilk boyutundaki elemanlar)
#print(len(satirlar))
jran = int((len(satirlar)-2)/3)	#boş satır sayısına göre kaç tablo olduğunu hesaplıyor
print(satirlar)

f1 = open("output.txt", "w")
#print (degerler_baslangic)


for j in range(0, jran+1):			#tablo döngüsü
	worksheet = wb.add_sheet('Sheet '+str(j+1))
	row = 0
	r1 = (satirlar[(j*3)]-(5*(j+1)))	#her tablonun başlangıç indisini hesaplıyor
		
	if j == jran:					#eğer tablo sayısı ile döngü indisi eşitse yani son tablodaysa
		r2 = degerler_sayisi			#tablonun sonuncu elemanı degerler dizisinin son elemanına eşittir
		buyukmu = (degerler_sayisi-(721))	#son tablonun sondan 720 elemanının olduğu indislerin minimumumu
	else:
		r2 = satirlar[(j*3+2)]-(j*3+2)			#eğer döngü son tabloda değilse tablonun sonuncu eleman indisini bul
		buyukmu = (satirlar[(j*3+2)]-(720+(j*3+3)))	#tablonun sondan 720 elemanının olduğu indislerin minimumumu
	
	mod = ((r2-r1)-(2*(j+3)))%720
	if mod != 0:
		f1.write("Bu tablonun mod 720 değeri = "+str(mod)+" r1="+str(r1+(2*(j+3)))+" r2="+str(r2)+"\n")
	clmn = 0
	for i in range(r1, r2):			#satır döngüsü (r1=tablonun degerler dizisindeki ilk elemanı, r2=vice versa)		

		if len(degerler[i]) == 3:	#eğer satırdaki ayrık eleman sayısı üç ise onlar başlıktır
			worksheet.write(row, clmn, str(degerler[i]).rstrip())
			clmn += 1
			f1.write(degerler[i][0]+" "+degerler[i][1]+" "+degerler[i][2]+"\t")	#başlığı yazdır
			continue					#diğer durumları atla

		if len(degerler[i]) == 1:	#eğer satırdaki ayrık eleman sayısı bir ise o gereksizdir
			row += 1
			f1.write("\n\n")	#iki satır aşağı in
			continue		#diğer durumları atla

		if i >= buyukmu:		#eğer satır indisi buyukmu sayısından büyükse satırları yazdır	
			row += 1		
			worksheet.write(row, 0, str(degerler[i][0]).rstrip())
			worksheet.write(row, 1, str(degerler[i][1]).rstrip())
			worksheet.write(row, 2, str(degerler[i][2]).rstrip())
			worksheet.write(row, 3, str(degerler[i][3]).rstrip())
			f1.write(degerler[i][0] +"\t"+ degerler[i][1] +"\t"+ degerler[i][2] +"\t"+ degerler[i][3] +"\n")
	
	f1.write("\n\n\n\n")			#tablo sonunda dört satır aşağı in
	
wb.save('output.xls')
f1.close()
lines.close()
