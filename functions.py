import matplotlib.pyplot as plt
from matplotlib.patches import Patch
import pyodbc as db
import numpy as np


# connection = db.connect('DRIVER={SQL Server};'
#                         'SERVER=-----------;'
#                         'DATABASE=ARCHIVESKF;'
#                         'UID=-----;PWD=------')
#  # 137.116.139.217
connection = db.connect('DRIVER={SQL Server};'
                        'SERVER=------------;'
                        'DATABASE=ARCHIVESKF;'
                        'UID=---;PWD=------')

cursor = connection.cursor()

query = """ 
        DECLARE @TotalItem VARCHAR(20)
        SET @TotalItem = (SELECT count( distinct  ITEMNO) as TotalMovingItem FROM ICStockStatusCurrentLOT 
                where  left(ITEMNO,1)<>'9' and AUDTORG<>'SKFDAT')
        SELECT NDMNAME, LEFT(AUDTORG,3) AS AUDTORG,
         ISNULL(Sum( CASE WHEN [Days]<=15 THEN 1 END),0) AS SUS
        ,ISNULL(Sum( CASE WHEN [Days]>15 AND [Days]<=35 THEN 1 END),0) AS US
        ,ISNULL(Sum( CASE WHEN [Days]>35 AND [Days]<=45 THEN 1 END),0) AS NS
        ,ISNULL(Sum( CASE WHEN [Days]>45 AND [Days]<=60 THEN 1 END),0) AS OS
        ,ISNULL(Sum( CASE WHEN [Days]>60 THEN 1 END),0) AS SOS
        ,@TotalItem-(ISNULL(Sum( CASE WHEN [Days]<=15 THEN 1 END),0)+ISNULL(Sum( CASE WHEN [Days]>15 AND [Days]<=35 THEN 1 END),0)+ISNULL(Sum( CASE WHEN [Days]>35 AND [Days]<=45 THEN 1 END),0)
        +ISNULL(Sum( CASE WHEN [Days]>45 AND [Days]<=60 THEN 1 END),0)+ISNULL(Sum( CASE WHEN [Days]>60 THEN 1 END),0)) AS NIL
        ,ISNULL(Sum( CASE WHEN [Days]>=0 THEN 1 END),0) AS TotalItem
        ,@TotalItem AS AllItem
        
        FROM
        
        (
        SELECT NDMNAME, Stock.ITEMNO, Stock.AUDTORG, Sum(QTYONHAND) as QTYONHAND , Sum(QTYSHIPPED)as QTYSHIPPED
        ,CAST(ISNULL(CASE WHEN SUM(QTYSHIPPED)>0 THEN (SUM(QTYONHAND)+ISNULL(SUM(SIT),0))/(SUM(QTYSHIPPED)/90) END,0) AS INT) AS [Days]FROM
        -----
        (SELECT ITEMNO, AUDTORG, SUM(QTYONHAND) AS QTYONHAND FROM ICStockStatusCurrentLOT
        where LEN(LOCATION)>'3' AND left(ITEMNO,1)<>'9' AND AUDTORG <> 'SKFDAT' 
        GROUP BY ITEMNO, AUDTORG) AS Stock
        LEFT JOIN
        (SELECT ITEM, AUDTORG, SUM(QTYSHIPPED) AS QTYSHIPPED FROM OESalesDetails
        WHERE TRANSDATE BETWEEN CONVERT(varchar(8), GETDATE()-91,112) AND CONVERT(varchar(8), GETDATE()-1,112)
        GROUP BY ITEM, AUDTORG) AS Sales
        ON RTRIM(Stock.ITEMNO) = RTRIM(Sales.ITEM) AND RTRIM(Stock.AUDTORG) = RTRIM(Sales.AUDTORG)
        LEFT JOIN
        (SELECT ITEMNO, AUDTORG,SUM(QTY) AS SIT FROM GIT WHERE OPENINGDATE = convert(varchar, getdate(), 23) GROUP BY ITEMNO, AUDTORG) as GIT
        ON Stock.ITEMNO = GIT.ITEMNO AND Stock.AUDTORG=GIT.AUDTORG
        left join
        (select distinct  BRANCHNAME,BRANCH,NDMNAME from NDM ) as NDM
        ON (Stock.AUDTORG=NDM.BRANCH)
        -------
        Group BY NDMNAME,Stock.ITEMNO, Stock.AUDTORG) AS TX
        GROUP BY NDMNAME,AUDTORG
        ORDER BY NDMNAME, AUDTORG
    """

cursor.execute(query)
data = list(cursor.fetchall())
print(data)
branchlist = []
SUS = []
US = []
NS = []
OS = []
SOS = []
NIL = []
BranchItem = []
TotalItem = []
rownumber = 0
for row in data:
        branchname = data[rownumber][1]
        branchlist.append(branchname)
        sus = data[rownumber][2]
        SUS.append(sus)
        us = data[rownumber][3]
        US.append(us)
        ns = data[rownumber][4]
        NS.append(ns)
        os = data[rownumber][5]
        OS.append(os)
        sos = data[rownumber][6]
        SOS.append(sos)
        nil = data[rownumber][7]
        NIL.append(nil)
        bi = data[rownumber][8]
        BranchItem.append(bi)
        ti = data[rownumber][9]
        TotalItem.append(ti)
        rownumber = rownumber+1



font = {'family': 'serif',
        'color':  '#3b5998',
        'weight': 400,
        'size': 15,
        }

font1 = {'family': 'serif',
        'color':  'darkred',
        'weight': 700,
        'size': 15,
        }


legend_elements = [Patch(facecolor='#be4748', label='Mr. Kamrul')
                   , Patch(facecolor='#1a58c5', label='Mr. Anwar')
                   , Patch(facecolor='#c5871a', label='Mr. Atik')
                   , Patch(facecolor='#58c51a', label='Mr. Nurul')
                   , Patch(facecolor='#fc0373', label='Mr. Hafizur')]

y_pos = np.arange(len(branchlist))

def bardecoration(bars):
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x(), yval + 3.5, yval, rotation=90, fontsize=8)

    for bar in bars[0:6]:
        bar.set_color('#be4748')

    for bar in bars[6:12]:
        bar.set_color('#1a58c5')

    for bar in bars[12:19]:
        bar.set_color('#c5871a')

    for bar in bars[19:25]:
        bar.set_color('#58c51a')

    for bar in bars[25:31]:
        bar.set_color('#fc0373')

    plt.xticks(y_pos, branchlist, rotation=90)
    plt.yticks(np.arange(0, 220, 20), fontsize=6)
    plt.ylabel('No. of Item', fontdict=font)
    plt.legend(handles=legend_elements, loc='best')

suskam = 0
uskam = 0
nskam = 0
soskam = 0
oskam = 0
nilkam = 0

for items in NIL[0:6]:
    nilkam += items

for items in SUS[0:6]:
    suskam += items

for items in US[0:6]:
    uskam += items

for items in SOS[0:6]:
    soskam += items

for items in OS[0:6]:
    oskam += items

for items in NS[0:6]:
    nskam += items

suswar = 0
uswar = 0
nswar = 0
soswar = 0
oswar = 0
nilwar = 0

for items in NIL[6:12]:
    nilwar += items

for items in SUS[6:12]:
    suswar += items

for items in US[6:12]:
    uswar += items

for items in SOS[6:12]:
    soswar += items

for items in OS[6:12]:
    oswar += items

for items in NS[6:12]:
    nswar += items

susbatik = 0
usbatik = 0
nsbatik = 0
sosbatik = 0
osbatik = 0
nilbatik = 0

for items in NIL[12:19]:
    nilbatik += items

for items in SUS[12:19]:
    susbatik += items

for items in US[12:19]:
    usbatik += items

for items in SOS[12:19]:
    sosbatik += items

for items in OS[12:19]:
    osbatik += items

for items in NS[12:19]:
    nsbatik += items

susnuru = 0
usnuru = 0
nsnuru = 0
sosnuru = 0
osnuru = 0
nilnuru = 0

for items in NIL[19:25]:
    nilnuru += items

for items in SUS[19:25]:
    susnuru += items

for items in US[19:25]:
    usnuru += items

for items in SOS[19:25]:
    sosnuru += items

for items in OS[19:25]:
    osnuru += items

for items in NS[19:25]:
    nsnuru += items

# ------------- For Hafizur
sushafizur = 0
ushafizur = 0
nshafizur = 0
soshafizur = 0
oshafizur = 0
nilhafizur = 0

for items in NIL[25:31]:
    nilhafizur += items

for items in SUS[25:31]:
    sushafizur += items

for items in US[25:31]:
    ushafizur += items

for items in SOS[25:31]:
    soshafizur += items

for items in OS[25:31]:
    oshafizur += items

for items in NS[25:31]:
    nshafizur += items


MovingItem = int(TotalItem[6])
print(MovingItem)

nuruta = [nilnuru/MovingItem, susnuru/MovingItem, usnuru/MovingItem, nsnuru/MovingItem, osnuru/MovingItem, sosnuru/MovingItem]
batikta = [nilbatik/MovingItem, susbatik/MovingItem, usbatik/MovingItem, nsbatik/MovingItem, osbatik/MovingItem, sosbatik/MovingItem]
warta = [nilwar/MovingItem, suswar/MovingItem, uswar/MovingItem, nswar/MovingItem, oswar/MovingItem, soswar/MovingItem]
kamta = [nilkam/MovingItem, suskam/MovingItem, uskam/MovingItem, nskam/MovingItem, oskam/MovingItem, soskam/MovingItem]
hafizur = [nilhafizur/MovingItem, sushafizur/MovingItem, ushafizur/MovingItem, nshafizur/MovingItem, soshafizur/MovingItem, oshafizur/MovingItem]

status = ['NIL', 'SUS', 'US', 'NS', 'OS', 'SOS']
y_pos1 = np.arange(len(status))

def branchbars(bars):
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x()+0.42, yval + 10.25, yval, horizontalalignment='center', verticalalignment='center', fontsize=14)

    plt.xticks(y_pos1, status, rotation=0)
    plt.yticks(np.arange(0, 220, 20), fontsize=10)
    plt.ylabel('No. of Item', fontdict=font)
    bars[0].set_color('#990000')
    bars[1].set_color('#ff0000')
    bars[2].set_color('#e6e600')
    bars[3].set_color('#008000')
    bars[4].set_color('#00ff16')
    bars[5].set_color('darkorange')


def sosbardecoration(bar1):
    for bar in bar1:
        yval = bar.get_height()
        plt.text(bar.get_x(), yval + 3.5, yval, rotation=90, fontsize = 8)

    for bar in bar1[0:6]:
        bar.set_color('#be4748')

    for bar in bar1[6:12]:
        bar.set_color('#1a58c5')

    for bar in bar1[12:19]:
        bar.set_color('#c5871a')

    for bar in bar1[19:25]:
        bar.set_color('#58c51a')

    for bar in bar1[25:31]:
        bar.set_color('#fc0373')

    plt.xticks(y_pos, branchlist, rotation=90)
    plt.yticks(np.arange(0, 440, 40), fontsize=6)
    plt.ylabel('No. of Item', fontdict=font)
    plt.legend(handles=legend_elements, loc='best')

print("Copleted")
