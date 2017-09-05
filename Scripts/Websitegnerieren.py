import xlszugriff

def Menu(fid_ziel,link,name):
    fid_ziel.write('\n<th class="Menu"><a href="' +link+ '.html" >' + name + '</a></th>')

def Menu1(fid_ziel,link):
    fid_ziel.write('\n<td class="Menu">'+ link +'</td>')

def Home(Textdatei,Zieldatei):
    fid_temp = open("Template.html","r")
    fid_ziel = open("..//"+Zieldatei,"w")
    fid_text = open(Textdatei,"r")
    if Zieldatei == 'index.html':
        for zeile in fid_temp:
            if zeile == "<!-- Inhalt1 -->\n":
                fid_ziel.write('\n<img class="Home" src="HomeBild1.jpg">\n')
            elif zeile == "<!-- Inhalt2 -->\n":
                fid_ziel.write('\n<ul>')
                for zeile in fid_text:
                    fid_ziel.write('\n<li>' + zeile + '</li>')
                fid_ziel.write('\n</ul>')
            else:
                fid_ziel.write(zeile)
        print Zieldatei + ":Fertig"
    elif Zieldatei == 'Impressum.html':
         for zeile in fid_temp:
            if zeile == "<!-- Inhalt1 -->\n":
                fid_ziel.write('\n<img class="Home" src="HomeBild1.jpg">\n')
            elif zeile == "<!-- Inhalt2 -->\n":
                fid_ziel.write('\n<ul>')
                for zeile in fid_text:
                    fid_ziel.write('\n<li>' + zeile + '</li>')
                fid_ziel.write('\n</ul>')
            else:
                fid_ziel.write(zeile)
         print Zieldatei + ":Fertig"
    elif Zieldatei == 'Gruppenphase.html':
        x = raw_input("Welche Gruppe spielt momentan?(A,B,C,D,E,F,G,H,K = Keine Gruppe spielt momentan): ")
        for zeile in fid_temp:
            if zeile == "<!-- Menu -->\n":
                fid_ziel.write('\n<table class="Menu">\n<thead>\n')
                for einzelLinie in daten:
                      Menu(fid_ziel,einzelLinie[0],einzelLinie[1])
                fid_ziel.write('\n</thead>\n<tbody>')
                for einzelLinie in daten:
                    Menu1(fid_ziel,einzelLinie[0])
                fid_ziel.write('\n</tbody>\n</table>')
            elif zeile == "<!-- Inhalt2 -->\n":
                for Datei in daten:
                    if x == Datei[1]:
                        y = Datei[2]
                        z = Datei[3]
                        fid_ziel.write('\n<table class="Tore">\n<thead>\n<tr>\n<th class="Tore">Spiel</th><th class="Tore">Datum</th><th class="Tore">Uhrzeit (MEZ)</th><th class="Tore">Austragungsort</th><th class="Tore" colspan="2">Mannschaften<th><th class="Tore" colspan="2">Tore</th>\n</tr>\n</thead>\n<tbody>\n')
                        Tabelle = xlszugriff.leseZellenbereich('WM-Tabelle.xlsx','Gruppenphase',y)
                        for Reihe in Tabelle:
                            fid_ziel.write('<tr>\n')
                            for Zelle in Reihe:
                                fid_ziel.write('<td class="Tore">%s</td>'%(Zelle))
                            fid_ziel.write('\n</tr>\n')
                        fid_ziel.write('</tbody>\n</table>\n')
                        fid_ziel.write('<table class="Punkte">\n<thead>\n<tr>\n<th class="Punkte">Platz</th><th class="Punkte">Mannschaft</th><th class="Punkte">Sp.</th><th class="Punkte">Pkte.</th><th class="Punkte" colspan="3">Tore</th><th class="Punkte">Diff.</th>\n</thead>\n<tbody>\n')
                        Tabelle1 = xlszugriff.leseZellenbereich('WM-Tabelle.xlsx','Gruppenphase',z)
                        for Reihe in Tabelle1:
                            fid_ziel.write('<tr>\n')
                            for Zelle in Reihe:
                                fid_ziel.write('<td class="Punkte">%s</td>'%(Zelle))
                            fid_ziel.write('\n</tr>\n')
                        fid_ziel.write('</tbody>\n</table>\n')
            elif zeile == "<!-- Inhalt1 -->\n":
                fid_ziel.write('\n<ul>')
                if x == "K":
                    for zeile in fid_text:
                        fid_ziel.write('\n<li>' + zeile + '</li>')
                    fid_ziel.write("\n<li>Momentan spielt keine Gruppe.</li>")
                else:
                    for zeile in fid_text:
                        fid_ziel.write('\n<li>' + zeile + '</li>')
                    fid_ziel.write("\n<li>Zurzeit spielt Gruppe " + x +": </li>")
                fid_ziel.write('\n</ul>')
            else:
                fid_ziel.write(zeile)
        print Zieldatei + ":Fertig"
    elif Zieldatei == 'KO-Phase.html':
        f = raw_input("Welches Finale ist momentan?(1/8,1/4,1/2,1,3,K = Kein Finale): ")
        for zeile in fid_temp:
            if zeile == "<!-- Menu -->\n":
                fid_ziel.write('\n<table class="Menu">\n<thead>\n')
                for einzelLinie in daten3:
                      Menu(fid_ziel,einzelLinie[0],einzelLinie[1])
                fid_ziel.write('\n</thead>\n<tbody>')
                for einzelLinie in daten3:
                    Menu1(fid_ziel,einzelLinie[0])
                fid_ziel.write('\n</tbody>\n</table>')  
            elif zeile == "<!-- Inhalt2 -->\n":
                for Datei in daten3:
                    if f == Datei[1]:
                       o = Datei[2]
                       fid_ziel.write('\n<table class="Finale">\n<thead>\n<tr>\n<th class="Finale">Spiel</th><th class="Finale">Datum</th><th class="Finale">Uhrzeit (MEZ)</th><th class="Finale">Austragungsort</th><th class="Finale" colspan="5">Tore<th>\n</tr>\n</thead>\n<tbody>\n')
                       Tabelle = xlszugriff.leseZellenbereich('WM-Tabelle.xlsx','KO-Phase',o)
                       for Reihe in Tabelle:
                            fid_ziel.write('<tr>\n')
                            for Zelle in Reihe:
                                fid_ziel.write('<td class="Finale">%s</td>'%(Zelle))
                            fid_ziel.write('\n</tr>\n')
                       fid_ziel.write('</tbody>\n</table>\n')
            elif zeile == "<!-- Inhalt1 -->\n":
                    fid_ziel.write('\n<ul>')
                    if f == "K":
                        for zeile in fid_text:
                            fid_ziel.write('\n<li>' + zeile + '</li>')
                        fid_ziel.write("\n<li>Momentan findet kein Finale statt</li>")
                    else:
                        for zeile in fid_text:
                            fid_ziel.write('\n<li>' + zeile + '</li>')
                        if f == "1":
                            fid_ziel.write("\n<li>Zurzeit ist das Finale:</li>")
                        else:
                            fid_ziel.write("\n<li>Zurzeit ist das " + f +" Finale: </li>")
                    fid_ziel.write('\n</ul>')
                
            else:
                fid_ziel.write(zeile)
    print Zieldatei + ":Fertig"


def Gruppen(Zieldatei,Daten):
    fid_temp = open("Template.html","r")
    fid_ziel = open("..//"+Zieldatei,"w")
    for zeile in fid_temp:
        if zeile == "<!-- Menu -->\n":
            fid_ziel.write('\n<table class="Menu">\n<thead>\n')
            for einzelLinie in daten:
                    Menu(fid_ziel,einzelLinie[0],einzelLinie[1])
            fid_ziel.write('\n</thead>\n<tbody>')
            for einzelLinie in daten:
                    Menu1(fid_ziel,einzelLinie[0])
            fid_ziel.write('\n</tbody>\n</table>')
        elif zeile == "<!-- Inhalt1 -->\n":
            fid_ziel.write('\n<table class="Tore">\n<thead>\n<tr>\n<th class="Tore">Spiel</th><th class="Tore">Datum</th><th class="Tore">Uhrzeit (MEZ)</th><th class="Tore">Austragungsort</th><th class="Tore" colspan="2">Mannschaften<th><th class="Tore" colspan="2">Tore</th>\n</tr>\n</thead>\n<tbody>\n')
            Tabelle = xlszugriff.leseZellenbereich('WM-Tabelle.xlsx','Gruppenphase',Datei[2])
            for Reihe in Tabelle:
                fid_ziel.write('<tr>\n')
                for Zelle in Reihe:
                    fid_ziel.write('<td class="Tore">%s</td>'%(Zelle))
                fid_ziel.write('\n</tr>\n')
            fid_ziel.write('</tbody>\n</table>\n')
        elif zeile == '<!-- Inhalt2 -->\n':
            fid_ziel.write('<table class="Punkte">\n<thead>\n<tr>\n<th class="Punkte">Platz</th><th class="Punkte">Mannschaft</th><th class="Punkte">Sp.</th><th class="Punkte">Pkte.</th><th class="Punkte" colspan="3">Tore</th><th class="Punkte">Diff.</th>\n</thead>\n<tbody>\n')
            Tabelle1 = xlszugriff.leseZellenbereich('WM-Tabelle.xlsx','Gruppenphase',Datei[3])
            for Reihe in Tabelle1:
                fid_ziel.write('<tr>\n')
                for Zelle in Reihe:
                    fid_ziel.write('<td class="Punkte">%s</td>'%(Zelle))
                fid_ziel.write('\n</tr>\n')
            fid_ziel.write('</tbody>\n</table>\n')
        else:
            fid_ziel.write(zeile)
    fid_temp.close()
    fid_ziel.close()
    print Zieldatei + ": Fertig!" 
    
def KOPhase(Zieldatei,Daten):
    fid_temp = open("Template.html","r")
    fid_ziel = open("..//"+Zieldatei,"w")
    for zeile in fid_temp:
        if zeile == "<!-- Menu -->\n":
            fid_ziel.write('\n<table class="Menu">\n<thead>\n')
            for einzelLinie in daten3:
                    Menu(fid_ziel,einzelLinie[0],einzelLinie[1])
            fid_ziel.write('\n</thead>\n<tbody>')
            for einzelLinie in daten3:
                    Menu1(fid_ziel,einzelLinie[0])
            fid_ziel.write('\n</tbody>\n</table>')
        elif zeile == "<!-- Inhalt1 -->\n":
            fid_ziel.write('\n<table class="Finale">\n<thead>\n<tr>\n<th class="Finale">Spiel</th><th class="Finale">Datum</th><th class="Finale">Uhrzeit (MEZ)</th><th class="Finale">Austragungsort</th><th class="Finale" colspan="5">Tore<th>\n</tr>\n</thead>\n<tbody>\n')
            Tabelle = xlszugriff.leseZellenbereich('WM-Tabelle.xlsx','KO-Phase',Datei[2])
            for Reihe in Tabelle:
                fid_ziel.write('<tr>\n')
                for Zelle in Reihe:
                    fid_ziel.write('<td class="Finale">%s</td>'%(Zelle))
                fid_ziel.write('\n</tr>\n')
            fid_ziel.write('</tbody>\n</table>\n')
        else:
            fid_ziel.write(zeile)
    fid_temp.close()
    fid_ziel.close()
    print Zieldatei + ": Fertig!" 
    
def Hochladen(Datei):
    session = ftplib.FTP("www5.subdomain.com","user2654671","abcd1234")
    session.cwd('www')
    
    HochladenDatei(session,Datei)
    session.close()
    
def HochladenDatei(session,Datei):
    fid = open("..//"+Datei,"r")
    session.storbinary("STOR "+Datei,fid)
    fid.close()
    
        
    

daten=[['Gruppe A','A','A8:J13','L9:S12'],['Gruppe B','B','A17:J22','L18:S21'],['Gruppe C','C','A26:J31','L27:S30'],['Gruppe D','D','A35:J40','L36:S39'],['Gruppe E','E','A44:J49','L45:S48'],['Gruppe F','F','A53:J58','L54:S57'],['Gruppe G','G','A62:J67','L63:S66'],['Gruppe H','H','A71:J76','L72:S75']]

daten2=[['index.html','Texte\\Home.txt'],['Gruppenphase.html','Texte\\Gruppenphase.txt'],['KO-Phase.html','Texte\\KO-Phase.txt'],['Impressum.html','Texte\\Impressum.txt']]

daten3=[['Achtelfinale','1/8','A9:I16'],['Viertelfinale','1/4','A21:I24'],['Halbfinale','1/2','A30:I31'],['Finale','1','A40:I40'],['Spiel um den dritten Platz','3','A36:I36']]

daten4=["Gruppe A.html","Gruppe B.html","Gruppe C.html","Gruppe D.html","Gruppe E.html","Gruppe F.html","Gruppe G.html","Gruppe H.html","index.html","Gruppenphase.html","KO-Phase.html","Impressum.html","Achtelfinale.html","Viertelfinale.html","Halbfinale.html","Finale.html","Spiel um den dritten Platz.html","WM2014.css"]
                        
for Datei in daten:
    Gruppen(Datei[0]+".html",daten)

for Datei in daten2:
    Home(Datei[1],Datei[0])

for Datei in daten3:
    KOPhase(Datei[0]+'.html',daten)

import ftplib

for Datei in daten4:
    Hochladen(Datei)
    print "Hochladen von " + Datei + " ist beendet"  

















