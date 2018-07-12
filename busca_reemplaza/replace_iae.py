fdisaeat = open ('DISTINTO_CODAEAT.txt', 'r') # fichero aeat
fread = open('MAT_DEF_IAE_2018.txt', 'r') #fichero de entrada
fwrite = open ('MAT_DEF_IAE_2018_mod.txt', 'w') #fichero de salida
#fread = open('pruebasrusti.txt', 'r')
#fwrite = open ('prusti2018_mod.txt', 'w')

#leemos cada linea del fichero. cogemos la cabecera para buscar en una zona de columna o otra.
#-----------------------------------------------
for linea in fread:   	

	cabecera = linea[0:1] 
	cod1 = linea[17:22]
	cod2 = linea[2:7]
	esta = 'bien'
		
#cad2 = line[405:425] 
	fdisaeat.seek(0)
	
	for line in fdisaeat:
		datomal = line[0:5]
		datobien = line[6:11]


		if cabecera == '1': 
						
			if cod1 == datomal: 
				fwrite.write(linea[:17] + datobien + linea[22:])
				esta = 'mal'
				continue
			else:
				continue
		
		elif cabecera == '3':
						
			if cod2 == datomal: 
				fwrite.write(linea[:2] + datobien + linea[7:])
				esta = 'mal'
				continue
			else:
				continue
						
		else:
			continue
	if esta == 'mal':
		continue
	else:
		fwrite.write(linea)
		continue
	


fread.close()
fwrite.close()
fdisaeat.close()