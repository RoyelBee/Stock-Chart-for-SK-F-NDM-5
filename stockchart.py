import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
import functions as fn
import linechart as ln
import os
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
from datetime import datetime
import time
from selenium import webdriver
from PIL import ImageEnhance
# from Screenshot_Clipping import Screenshot
dirpath = os.path.dirname(os.path.realpath(__file__))


# try:
#     os.remove(dirpath+'/SKFStockReport.xlsx')
#     print("File Removed!")
# except WindowsError:
#     print("no file found")
#     pass

print('Execution start')
plt.figure(0)
bars = plt.bar(fn.y_pos, fn.NIL, width=0.5)
fn.bardecoration(bars)
plt.title('Branch Status - Stock NIL', fontdict=fn.font1)
plt.savefig(dirpath+'/nilbar.png')

plt.figure(1)
bars = plt.bar(fn.y_pos, fn.SUS, width=0.5)
fn.bardecoration(bars)
plt.title('Branch Status - Stock SUS (1-15 Days)', fontdict=fn.font1)
plt.savefig(dirpath+'/susbar.png')

plt.figure(2)
bars = plt.bar(fn.y_pos, fn.US, width=0.5)
fn.bardecoration(bars)
plt.title('Branch Status - Stock US (16-35 Days)', fontdict=fn.font1)
plt.savefig(dirpath+'/usbar.png')

plt.figure(3)
bars = plt.bar(fn.y_pos, fn.NS, width=0.5)
fn.bardecoration(bars)
plt.title('Branch Status - Stock NS (36-45 Days)', fontdict=fn.font1)
plt.savefig(dirpath+'/nsbar.png')

plt.figure(4)
bars = plt.bar(fn.y_pos, fn.OS, width=0.5)
fn.bardecoration(bars)
plt.title('Branch Status - Stock OS (46-60 Days)', fontdict=fn.font1)
plt.savefig(dirpath+'/osbar.png')

plt.figure(5)
bar1 = plt.bar(fn.y_pos, fn.SOS, width=0.5)
fn.sosbardecoration(bar1)
plt.title('Branch Status - Stock SOS (More Than 60 Days)', fontdict=fn.font1)
plt.savefig(dirpath+'/sosbar.png')

image = Image.new('RGB', (1281, 1442))
im = Image.open(dirpath+"/nilbar.png")
width,height=im.size
im1 = Image.open(dirpath+"/susbar.png")
im2 = Image.open(dirpath+"/usbar.png")
im3 = Image.open(dirpath+"/sosbar.png")
im4 = Image.open(dirpath+"/osbar.png")
im5 = Image.open(dirpath+"/nsbar.png")


#### image = Image.new('RGB', (width1+width2+width3+width4, height+height1+1))
# Set image position in the mail body
image.paste(im, (0, 0))
image.paste(im5, (width+1, 0))

image.paste(im2, (0, height+1))
image.paste(im1, (width+1, height+1))

image.paste(im4, (0, height*2+2))
image.paste(im3, (width+1, height*2+2))

image.save(dirpath+"/Final.png")


# -----------------------------------------------------------------------
# -------------------------Pie Chart start-------------------------------
# -----------------------------------------------------------------------

colors = ['#990000', '#ff0000', '#e6e600', '#008000', '#00ff16', 'darkorange']
status = ['NIL', 'SUS', 'US', 'NS', 'OS', 'SOS']

plt.figure(6)
wtf = plt.pie(fn.nuruta, colors=colors, autopct='%1.1f%%', pctdistance=1.2)
plt.title('Mr. Nurul', fontdict=fn.font1)
plt.legend(labels=status, title="Status", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))

fraction_text_list = wtf[2]
for text in fraction_text_list:
    text.set_rotation(70)
plt.savefig(dirpath+'/nurupie.png')

plt.figure(7)
wtf = plt.pie(fn.batikta, colors=colors, autopct='%1.1f%%', pctdistance=1.2)
plt.title('Mr. Atik', fontdict=fn.font1)
plt.legend(labels=status, title="Status", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))

fraction_text_list = wtf[2]
for text in fraction_text_list:
    text.set_rotation(70)
plt.savefig(dirpath+'/batikpie.png')

plt.figure(8)
wtf = plt.pie(fn.warta, colors=colors, autopct='%1.1f%%', pctdistance=1.2)
plt.title('Mr. Anwar', fontdict=fn.font1)
plt.legend(labels=status, title="Status", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))

fraction_text_list = wtf[2]
for text in fraction_text_list:
    text.set_rotation(70)
plt.savefig(dirpath+'/warpie.png')

plt.figure(9)
wtf = plt.pie(fn.kamta, colors=colors, autopct='%1.1f%%', pctdistance=1.2)
plt.title('Mr. Kamrul', fontdict=fn.font1)
plt.legend(labels=status, title="Status", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))

fraction_text_list = wtf[2]
for text in fraction_text_list:
    text.set_rotation(70)
plt.savefig(dirpath+'/kampie.png')

plt.figure(10)
wtf = plt.pie(fn.kamta, colors=colors, autopct='%1.1f%%', pctdistance=1.2)
plt.title('Mr. Hafizur', fontdict=fn.font1)
plt.legend(labels=status, title="Status", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))

fraction_text_list = wtf[2]
for text in fraction_text_list:
    text.set_rotation(70)
plt.savefig(dirpath+'/hafizurpie.png')



imageX = Image.new('RGB', (1280, 1440))

imp1 = Image.open(dirpath+"/kampie.png")
widthx,heightx = imp1.size
imp2 = Image.open(dirpath+"/warpie.png")
imp3 = Image.open(dirpath+"/batikpie.png")
imp4 = Image.open(dirpath+"/nurupie.png")
imp5 = Image.open(dirpath+"/hafizurpie.png")
imp6 = Image.open(dirpath+"/fakebar.png")


imageX.paste(imp1, (0, 0))
imageX.paste(imp2, (widthx, 0))
imageX.paste(imp3, (0, heightx))
imageX.paste(imp4, (widthx, heightx))
imageX.paste(imp5, (0, height*2))
imageX.paste(imp6, (widthx, height*2))


imageX.save(dirpath+"/ndmpie.png")
#
# # ------------------------------------------------
# # --------------NDM bar chart start ------------
# # ------------------------------------------------

i = 0
x = 11
y_pos = np.arange(len(status))
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('CTG Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/ctgbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('CTN Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/ctnbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('NAJ Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/najbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('COX Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/coxbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('KUS Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/kusbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('PBN Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/pbnbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('JES Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/jesbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('MOT Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/motbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('MYM Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/mymbar.png')



i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('NOK Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/nokbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('COM Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/combar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('FEN Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/fenbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('FRD Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/frdbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('GZP Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/gzpbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('KHL Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/khlbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('KSG Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/ksgbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('MIR Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/mirbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('PAT Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/patbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('BOG Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/bogbar.png')



i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('HZJ Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/hzjbar.png')


i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('MHK Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/mhkbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('MLV Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/mlvbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('SYL Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/sylbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('BSL Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/bslbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('DNJ Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/dnjbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('KRN Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/krnbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('RAJ Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/rajbar.png')

i = i+1
x = x+1
plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('RNG Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/rngbar.png')
i = i+1
x = x+1

plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('SAV Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/savbar.png')
i = i+1
x = x+1

plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('TGL Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/tglbar.png')
i = i+1
x = x+1

plt.figure(x)
mapdata = [fn.NIL[i], fn.SUS[i], fn.US[i], fn.NS[i], fn.OS[i], fn.SOS[i]]
Bitem = fn.BranchItem[i]
TItem = fn.TotalItem[i]
bars = plt.bar(y_pos, mapdata, align='center', alpha=1)
fn.branchbars(bars)
plt.title('VRB Branch - Stock Count('+str(Bitem)+' of '+str(TItem)+')', fontdict=fn.font1)
plt.savefig(dirpath+'/vrbbar.png')
i = i+1
x = x+1

image = Image.new('RGB', (width*3, height*3))
im1 = Image.open(dirpath+"/hzjbar.png")
width, height = im1.size
im2 = Image.open(dirpath+"/motbar.png")
im3 = Image.open(dirpath+"/ksgbar.png")
im4 = Image.open(dirpath+"/gzpbar.png")
im5 = Image.open(dirpath+"/rngbar.png")
im6 = Image.open(dirpath+"/dnjbar.png")
im7 = Image.open(dirpath+"/krnbar.png")
imfk = Image.open(dirpath+"/fakebar.png")
imtik = Image.open(dirpath+"/batik.png")


image.paste(imtik, (0, 0))
image.paste(im1, (width, 0))
image.paste(im2, (width*2, 0))
image.paste(im3, (0, height))
image.paste(im4, (width, height))
image.paste(im5, (width*2, height))
image.paste(im6, (0, height*2))
image.paste(im7, (width, height*2))
image.paste(imfk, (width*2, height*2))

image.save(dirpath+"/ndm1.png")

# - kamrul bar chart start ------------------
image = Image.new('RGB', (width*3, height*3))
im1 = Image.open(dirpath+"/jesbar.png")
width, height = im1.size
im2 = Image.open(dirpath+"/mirbar.png")
im3 = Image.open(dirpath+"/khlbar.png")
im4 = Image.open(dirpath+"/combar.png")
im5 = Image.open(dirpath+"/patbar.png")
im6 = Image.open(dirpath+"/bslbar.png")
kamfk = Image.open(dirpath+"/fakebar.png")
imkam = Image.open(dirpath+"/kam.png")


image.paste(imkam, (0, 0))
image.paste(im1, (width, 0))
image.paste(im2, (width*2, 0))
image.paste(im3, (0, height))
image.paste(im4, (width, height))
image.paste(im5, (width*2, height))
image.paste(im6, (0, height*2))
image.paste(kamfk, (width, height*2))
image.paste(kamfk, (width*2, height*2))

image.save(dirpath+"/ndm2.png")
# - kamrul bar chart end ------------------


# - Start barchart for Md. Nurul Amin ------
image = Image.new('RGB', (width*3, height*3))
im1 = Image.open(dirpath+"/mlvbar.png")
width, height = im1.size
im2 = Image.open(dirpath+"/mhkbar.png")
im3 = Image.open(dirpath+"/sylbar.png")
im4 = Image.open(dirpath+"/nokbar.png")
im5 = Image.open(dirpath+"/fenbar.png")
im6 = Image.open(dirpath+"/vrbbar.png")
nurulfk = Image.open(dirpath+"/fakebar.png")
imnurul = Image.open(dirpath+"/nuru.png")

image.paste(imnurul, (0, 0))
image.paste(im1, (width, 0))
image.paste(im2, (width*2, 0))
image.paste(im3, (0, height))
image.paste(im4, (width, height))
image.paste(im5, (width*2, height))
image.paste(im6, (0, height*2))
image.paste(nurulfk, (width, height*2))
image.paste(nurulfk, (width*2, height*2))

image.save(dirpath+"/ndm3.png")
# - End bar chart for Md. Nurul Amin ------


# - Start bar chart for Md. Anwar ------
image = Image.new('RGB', (width*3, height*3))
im1 = Image.open(dirpath+"/bogbar.png")
width, height = im1.size
im2 = Image.open(dirpath+"/mymbar.png")
im3 = Image.open(dirpath+"/frdbar.png")
im4 = Image.open(dirpath+"/tglbar.png")
im5 = Image.open(dirpath+"/rajbar.png")
im6 = Image.open(dirpath+"/savbar.png")
anwarfk = Image.open(dirpath+"/fakebar.png")
imanwar = Image.open(dirpath+"/war.png")

image.paste(imanwar, (0, 0))
image.paste(im1, (width, 0))
image.paste(im2, (width*2, 0))
image.paste(im3, (0, height))
image.paste(im4, (width, height))
image.paste(im5, (width*2, height))
image.paste(im6, (0, height*2))
image.paste(anwarfk, (width, height*2))
image.paste(anwarfk, (width*2, height*2))

image.save(dirpath+"/ndm4.png")
# - End bar chart for Anwar  ------

# - Start barch art for Hafizur ------
image = Image.new('RGB', (width*3, height*3))
im1 = Image.open(dirpath+"/kusbar.png")
width, height = im1.size
im2 = Image.open(dirpath+"/pbnbar.png")
im3 = Image.open(dirpath+"/coxbar.png")
im4 = Image.open(dirpath+"/najbar.png")
im5 = Image.open(dirpath+"/ctgbar.png")
im6 = Image.open(dirpath+"/ctnbar.png")
hafizurfk = Image.open(dirpath+"/fakebar.png")
imhafizur = Image.open(dirpath+"/hafizur.png")

image.paste(imhafizur, (0, 0))
image.paste(im1, (width, 0))
image.paste(im2, (width*2, 0))
image.paste(im3, (0, height))
image.paste(im4, (width, height))
image.paste(im5, (width*2, height))
image.paste(im6, (0, height*2))
image.paste(hafizurfk, (width, height*2))
image.paste(hafizurfk, (width*2, height*2))

image.save(dirpath+"/ndm5.png")
# - End bar chart for Hafizur  ------
# # -----------------------------------------
# # NDM Bar Chart End
# # -----------------------------------------
#
#
# # -----------------------------------------
# # --------- Line Chart Start --------------
# # -----------------------------------------
plt.figure(x)
plt.plot(ln.days, ln.nilkam30, color='#be4748')
plt.plot(ln.days, ln.nilwar30, color='#1a58c5')
plt.plot(ln.days, ln.niltik30, color='#c5871a')
plt.plot(ln.days, ln.nilnuru30, color='#58c51a')
plt.plot(ln.days, ln.nilfiz30, color='#fc0373')

plt.xticks(range(1,31), ln.datenuru30, rotation=90)
plt.ylabel('No. of Item', fontdict=ln.font)
plt.xlabel('Days', fontdict=ln.font)
plt.title('NDM\'s Last 30 Days NIL Status', fontdict=ln.font1)
plt.legend(['Mr. Kamrul', 'Mr. Anwar', 'Mr. Atik', 'Mr. Nurul', 'Mr. Hafizur'], loc='best')
plt.savefig(dirpath+'/nilline.png')
x=x+1

plt.figure(x)
plt.plot(ln.days, ln.uskam30, color='#be4748')
plt.plot(ln.days, ln.uswar30, color='#1a58c5')
plt.plot(ln.days, ln.ustik30, color='#c5871a')
plt.plot(ln.days, ln.usnuru30, color='#58c51a')
plt.plot(ln.days, ln.usfiz30, color='#fc0373')

plt.xticks(range(1,31), ln.datenuru30, rotation=90)
plt.ylabel('No. of Item', fontdict=ln.font)
plt.xlabel('Days', fontdict=ln.font)
plt.title('NDM\'s Last 30 Days US Status', fontdict=ln.font1)
plt.legend(['Mr. Kamrul', 'Mr. Anwar', 'Mr. Atik', 'Mr. Nurul', 'Mr. Hafizur'], loc='best')
plt.savefig(dirpath+'/usline.png')
x=x+1

plt.figure(x)
plt.plot(ln.days, ln.suskam30, color='#be4748')
plt.plot(ln.days, ln.suswar30, color='#1a58c5')
plt.plot(ln.days, ln.sustik30, color='#c5871a')
plt.plot(ln.days, ln.susnuru30, color='#58c51a')
plt.plot(ln.days, ln.susfiz30, color='#fc0373')

plt.xticks(range(1,31), ln.datenuru30, rotation=90)
plt.ylabel('No. of Item', fontdict=ln.font)
plt.xlabel('Days', fontdict=ln.font)
plt.title('NDM\'s Last 30 Days SUS Statsus', fontdict=ln.font1)
plt.legend(['Mr. Kamrul', 'Mr. Anwar', 'Mr. Atik', 'Mr. Nurul', 'Mr. Hafizur'], loc='best')
plt.savefig(dirpath+'/susline.png')
x=x+1

plt.figure(x)
plt.plot(ln.days, ln.nskam30, color='#be4748')
plt.plot(ln.days, ln.nswar30, color='#1a58c5')
plt.plot(ln.days, ln.nstik30, color='#c5871a')
plt.plot(ln.days, ln.nsnuru30, color='#58c51a')
plt.plot(ln.days, ln.nsfiz30, color='#fc0373')

plt.xticks(range(1,31), ln.datenuru30, rotation=90)
plt.ylabel('No. of Item', fontdict=ln.font)
plt.xlabel('Days', fontdict=ln.font)
plt.title('NDM\'s Last 30 Days NS Status', fontdict=ln.font1)
plt.legend(['Mr. Kamrul', 'Mr. Anwar', 'Mr. Atik', 'Mr. Nurul', 'Mr. Hafizur'], loc='best')
plt.savefig(dirpath+'/nsline.png')
x=x+1

plt.figure(x)
plt.plot(ln.days, ln.oskam30, color='#be4748')
plt.plot(ln.days, ln.oswar30, color='#1a58c5')
plt.plot(ln.days, ln.ostik30, color='#c5871a')
plt.plot(ln.days, ln.osnuru30, color='#58c51a')
plt.plot(ln.days, ln.osfiz30, color='#fc0373')

plt.xticks(range(1,31), ln.datenuru30, rotation=90)
plt.ylabel('No. of Item', fontdict=ln.font)
plt.xlabel('Days', fontdict=ln.font)
plt.title('NDM\'s Last 30 Days OS Status', fontdict=ln.font1)
plt.legend(['Mr. Kamrul', 'Mr. Anwar', 'Mr. Atik', 'Mr. Nurul', 'Mr. Hafizur'], loc='best')
plt.savefig(dirpath+'/osline.png')
x = x + 1

plt.figure(x)
plt.plot(ln.days, ln.soskam30, color='#be4748')
plt.plot(ln.days, ln.soswar30, color='#1a58c5')
plt.plot(ln.days, ln.sostik30, color='#c5871a')
plt.plot(ln.days, ln.sosnuru30, color='#58c51a')
plt.plot(ln.days, ln.sosfiz30, color='#fc0373')

plt.xticks(range(1, 31), ln.datenuru30, rotation=90)
plt.ylabel('No. of Item', fontdict=ln.font)
plt.xlabel('Days', fontdict=ln.font)
plt.title('NDM\'s Last 30 Days SOS Status', fontdict=ln.font1)
plt.legend(['Mr. Kamrul', 'Mr. Anwar', 'Mr. Atik', 'Mr. Nurul', 'Mr. Hafizur'], loc='best')
plt.savefig(dirpath+'/sosline.png')
x=x+1


imageline = Image.new('RGB', (1281, 1442))

imp1  = Image.open(dirpath+"/nilline.png")
widthx,heightx = imp1.size
imp2  = Image.open(dirpath+"/nsline.png")
imp3  = Image.open(dirpath+"/usline.png")
imp4  = Image.open(dirpath+"/susline.png")
imp5  = Image.open(dirpath+"/osline.png")
imp6  = Image.open(dirpath+"/sosline.png")


imageline.paste(imp1,(0,0))
imageline.paste(imp2,(widthx+1,0))
imageline.paste(imp3,(0,heightx+1))
imageline.paste(imp4,(widthx+1,heightx+1))
imageline.paste(imp5,(0,heightx*2+2))
imageline.paste(imp6,(widthx+1,heightx*2+2))

imageline.save(dirpath+"/ndmline.png")

# ------------------------------------------------
# ---- Line Chart End ----------------------------
# ------------------------------------------------

########---------Wbdriver Part


options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
prefs = {"profile.default_content_settings.popups": 0,"download.default_directory": dirpath,"directory_upgrade": True, "safebrowsing.enabled": True}
options.add_experimental_option("prefs", prefs)

driver=webdriver.Chrome('chromedriver.exe', options=options)
driver.set_page_load_timeout(1500)

#ob=Screenshot_Clipping.Screenshot()
driver.get("http://Royel:ro@123@skfonline.transcombd.com:9081/reports/report/PHARMA/RnD%20(Only%20ERP)/Mail%20Automation%20Report/sum")
time.sleep(40)
driver.save_screenshot('dhd.png')

im = Image.open(dirpath+"/dhd.png")
croppedIm = im.crop((0, 43, 1010, 580))

# enhancer = ImageEnhance.Sharpness(croppedIm)
# enhanced_im = enhancer.enhance(0.8)

croppedIm.save(dirpath+"/dhd_col.png", "png", quality=100, optimize=True, progressive=True)

# driver.get('http://Alim:alimerp@10.168.2.245:9081/reports/')
# print("Downloading Stock Report... ...")
# driver.get("http://Alim:alimerp@10.168.2.245:9081/Report?%2FSKFStockReport&rs:Format=EXCELOPENXML")
# print("Download Complete! Make Happy Face :D ")

driver.get('http://Royel:ro@123@skfonline.transcombd.com:9081/reports/browse/PHARMA/RnD%20(Only%20ERP)/Mail%20Automation%20Report')
print("Downloading Stock Report... ...")
driver.get("http://Royel:ro@123@skfonline.transcombd.com:9081/reports/report/PHARMA/RnD%20(Only%20ERP)/Mail%20Automation%20Report/SKFStockReport:Format=EXCELOPENXML")
print("Download Complete! Make Happy Face :D ")

time.sleep(17)
driver.close()


msgRoot = MIMEMultipart('related')
me = 'erp-bi.service@transcombd.com'
to = ['hislam@skf.transcombd.com', 'muhammad.mainuddin@tdcl.transcombd.com']
cc = ['almamun@transcombd.com', 'tdclndm@tdcl.transcombd.com', 'tdclpharma@transcombd.com', 'monowar@tdcl.transcombd.com', 'mosaddek.hossain@skf.transcombd.com', 'monirul@skf.transcombd.com']
bcc = ['biswascma@yahoo.com', 'bayezid@transcombd.com', 'zubair.transcom@gmail.com', 'yakub@transcombd.com', 'redwan@transcombd.com', 'tawhid@transcombd.com', 'rejaul.islam@transcombd.com']
recipient = to+cc+bcc

# to = 'rejaul.islam@transcombd.com'
# recipient = to

date = datetime.today()
today = str(date.day)+'/'+str(date.month)+'/'+str(date.year)+' '+date.strftime("%I:%M %p")
today1 = str(date.day)+' '+str(date.strftime("%B"))+' '+str(date.year)+' at '+date.strftime("%I:%M %p")

subject = "SKF Formulation – Stock Status Report - "+today

email_server_host = 'mail.transcombd.com'
port = 25

msgAlternative = MIMEMultipart('alternative')
msgRoot.attach(msgAlternative)

msgText = MIMEText('This is the alternative plain text message.')
msgAlternative.attach(msgText)

# We reference the image in the IMG SRC attribute by the ID we give it below
msgText = MIMEText("""<b>Dear Sir,</b><br><br>Enclosed, Please find herewith the Stock Status report of """
                   + today1 + """ for TDCL All Branches including Central Warehouse.<br><br>

                   <img src="cid:image1"><br> <br>
                    <img src="cid:image2"><br> <br>
                   <img src="cid:image8"><br> <br>
                   <img src="cid:image3"><br> <br>
                   <img src="cid:image4"><br> <br>
                   <img src="cid:image5"><br> <br>
                   <img src="cid:image6"><br> <br>
                   <img src="cid:image7"><br> <br>
                   <img src="cid:image9"><br> <br>

                   If there is any inconvenience, you are requested to communicate with the ERP BI Service:
                   <br><b>(Mobile: 01713-389972, 01713-380499)</b><br><br>
                   Regards<br><b>ERP BI Service</b><br>Information System Automation (ISA)<br><br>
                   <i><font color="blue">****This is a system generated stock report ****</i></font>""", 'html')


msgAlternative.attach(msgText)

# Assigning the 1st image (NDM Bar Chart)
fp = open(dirpath+'/Final.png', 'rb')
msgImage1 = MIMEImage(fp.read())
fp.close()

msgImage1.add_header('Content-ID', '<image1>')
msgRoot.attach(msgImage1)


# NDM pie chart are stored in image2
fp = open(dirpath+'/ndmpie.png', 'rb')
msgImage2 = MIMEImage(fp.read())
fp.close()

# Define the 2nd image's ID as referenced above
msgImage2.add_header('Content-ID', '<image2>')
msgRoot.attach(msgImage2)

#
# # NDM bar chart reference
# # ############################
fp = open(dirpath+'/ndm1.png', 'rb')
msgImage3 = MIMEImage(fp.read())
fp.close()

# Define the 2nd image's ID as referenced above
msgImage3.add_header('Content-ID', '<image3>')
msgRoot.attach(msgImage3)

# Assigning the 2nd image directory
fp = open(dirpath+'/ndm2.png', 'rb')
msgImage4 = MIMEImage(fp.read())
fp.close()

# Define the 2nd image's ID as referenced above
msgImage4.add_header('Content-ID', '<image4>')
msgRoot.attach(msgImage4)

# Assigning the 2nd image directory
fp = open(dirpath+'/ndm3.png', 'rb')
msgImage5 = MIMEImage(fp.read())
fp.close()

# Define the 2nd image's ID as referenced above
msgImage5.add_header('Content-ID', '<image5>')
msgRoot.attach(msgImage5)

# Assigning the 2nd image directory
fp = open(dirpath+'/ndm4.png', 'rb')
msgImage6 = MIMEImage(fp.read())
fp.close()

# Define the 2nd image's ID as referenced above
msgImage6.add_header('Content-ID', '<image6>')
msgRoot.attach(msgImage6)

# Assigning the 2nd image directory
fp = open(dirpath+'/ndm5.png', 'rb')
msgImage7 = MIMEImage(fp.read())
fp.close()
# Define the 2nd image's ID as referenced above
msgImage7.add_header('Content-ID', '<image7>')
msgRoot.attach(msgImage7)


# ------ Line chart
# ---- Assign image directory
fp = open(dirpath+'/ndmline.png', 'rb')
msgImage8 = MIMEImage(fp.read())
fp.close()

msgImage8.add_header('Content-ID', '<image8>')
msgRoot.attach(msgImage8)
#
fp = open(dirpath+'/dhd_col.png', 'rb')
msgImage9 = MIMEImage(fp.read())
fp.close()

msgImage9.add_header('Content-ID', '<image9>')
msgRoot.attach(msgImage9)


#Excel attachment
part = MIMEBase('application', "octet-stream")
file_location = dirpath+'/SKFStockReport.xlsx'
###print(file_location)
import os
# Create the attachment file (only do it once)
filename = os.path.basename(file_location)
attachment = open(file_location, "rb")
part = MIMEBase('application', 'octet-stream')
part.set_payload(attachment.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
msgRoot.attach(part)


## For multiple users
msgRoot['From'] = me
msgRoot['To'] = ', '.join(to)
msgRoot['Cc'] = ', '.join(cc)
msgRoot['Bcc'] = ', '.join(bcc)
msgRoot['Subject'] = subject

## For single user
# msgRoot['to'] = recipient
# msgRoot['from'] = me
# msgRoot['subject'] = subject

server = smtplib.SMTP(email_server_host, port)
server.ehlo()
print('Sending Mail')
server.sendmail(me, recipient, msgRoot.as_string())
print('Mail Sent')
server.close()
