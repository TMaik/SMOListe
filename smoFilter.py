import xlsxwriter
import time
import os
import glob
import csv
# ******************************************************************************
#
# Deklarationen
#
# ******************************************************************************

# Anzahl der zu behaltenden alten Dateien
amountOfOldFiles = 2

# Aktuelle Länge der Dateiliste
listLength = 0


# ******************************************************************************
#
# Programmablauf
#
# ******************************************************************************

# Lösche die alten Listen
listFiles = glob.glob( './SMO Filterliste*.xlsx' )

if listLength > 3 :
        # sortieren
        listFiles.sort( reverse = True )        
        # löschen
        for i in listFiles[ amountOfOldFiles: ] :
                os.remove( i )

# Erzeuge neue Liste
strFileName = time.strftime( '%Y-%m-%d %H.%M' )
workbook = xlsxwriter.Workbook( 'SMO Filterliste ' + strFileName + '.xlsx' )

# Öffne und filter SMO Datei
reader = csv.reader( open( 'SMO_NB_REPORT_22-11-2018-01-00.csv', 'r' ),delimiter = ';' )

# Kopiere Überschriften aus csv Datei in Tabellenblatt
wsMaik	= workbook.add_worksheet( 'Maik' )
print( 'erzeuge wsMaik' )
for r in reader: 
        for c in r:
                if r.index() > 0: # TODO: iterieren durch liste oder reader
                        break
                wsMaik.write( c )

filtered = list(
                # Rufnummer = [16] startet mit FXP
                filter( lambda p: p[16].strip().startswith( 'FXP' ), reader
                )
           )



for i in filtered[ 0 : 1 ] :
        print( i )


wsJutta = workbook.add_worksheet( 'Jutta' )
wsMatse = workbook.add_worksheet( 'Matse' )

workbook.close()
