# -*- encoding: utf-8 -*-

import xlrd,datetime

def leseZelle(book,sheet,rowx,colx):
    sh= book.sheet_by_name(sheet)
    #print rowx,colx
    cval = sh.cell_value(rowx, colx)
    cty = sh.cell_type(rowx, colx)
    #print "val",cty,">",cval,"<",rowx, colx
    if (cty == xlrd.XL_CELL_TEXT or cty == xlrd.XL_CELL_BOOLEAN):
        cval = cval.replace(" 2014","")
        cval = cval.replace(". ",".")
        return cval
    elif (cty == xlrd.XL_CELL_DATE):
        datevalue = xlrd.xldate_as_tuple(cval, book.datemode)
        if (cval < 1.0):
            return datetime.time(hour = datevalue[3], minute = datevalue[4]).strftime('%H:%M')
        else:
            return datetime.datetime(year = datevalue[0], month = datevalue[1], day = datevalue[2], hour = datevalue[3], minute = datevalue[4], second = datevalue[5]).strftime('%d.%B.%Y %H:%M')
    elif (cty == xlrd.XL_CELL_NUMBER):
        if cval.as_integer_ratio()[1]==1:
            return int(cval)
        else:
            return cval
    elif (cty == xlrd.XL_CELL_EMPTY):
        return ""
    else:
        print "error on:",xrow,xcol,cty
        raise Exception

#print dir(book)

def leseZellenbereich(xlsfile,sheet,cellrange):
    """
Diese Funktion liest einen Zellenbereich aus einer EXCEL-Datei.

Parameter:
    xlsfile (Zeichenkette):   Der Name der EXCEL-Datei
    sheet (Zeichenkette):     Der Name des Sheets in der EXCEL-Datei
    cellrange (Zeichenkette): Der Zellenbereich im bekannten EXCEL-Format (z.B.
                              "A4:C8")

Rückgabewert:
    Eine verschachtelte Liste der ausgewählten Zellen.  Jede Zeile ist
    eine Liste der Zellen in dieser Zeile.  Alle Zeilen werden in einer Liste
    geordnet aufgelistet.

Beispiel:
    leseZellenbereich( "Beispiel.xlsx" , "Tabelle1" , "C2:D5" )
    -> gibt den Zellenbereich "C2:D5" aus dem Sheet "Tabelle1" der EXCEL-Tabelle
       "Beispiel.xlsx" zurück.
"""
    book = xlrd.open_workbook(xlsfile)
    cells = cellrange.split(":")
    if (len(cells)==2):
        startcell = cells[0]
        endcell = cells[1]
    elif len(cells)==1:
        startcell = cells[0]
        endcell = cells[0]
    else:
        raise Exception("Falscher Zellenbereich " + cellrange )
        
    colstart = ord(startcell[0])-ord("A")
    rowstart = int(startcell[1:])-1
    colend = ord(endcell[0])-ord("A")+1
    rowend = int(endcell[1:])
    celllist = []
    print "leseZellenbereich",xlsfile,sheet,startcell, endcell
    print startcell,"->",rowstart,colstart
    print endcell,"->",rowend,colend
    for row in range(rowstart,rowend):
        #print "row",row
        rowlist = []
        for col in range(colstart,colend):
            #print "row/col",row,col
            #print leseZelle(book,sheet,row,col),":::",
            rowlist.append(leseZelle(book,sheet,row,col))
        #print ""
        celllist.append(rowlist)
    book = None
    #print celllist
    return celllist
