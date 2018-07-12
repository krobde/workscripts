#programa para reemplazo masivo buscando en una zona de columnas 
#dependiendo de la cabecera busca en una zona u otra

#import pdb; 
#pdb.set_trace()


fread = open('RUSTICA_2018.txt', 'r') #fichero de entrada
fwrite = open ('RUSTICA_2018_mod.txt', 'w') #fichero de salida
#fread = open('pruebasrusti.txt', 'r')
#fwrite = open ('prusti2018_mod.txt', 'w')

#leemos cada linea del fichero. cogemos la cabecera para buscar en una zona de columna o otra.
for line in fread:   	
	cabecera = line[:2] 
	cad1 = line[802:822] 
	cad2 = line[405:425] 
	
	# si la cabecera es 54 buscamos en los caracteres desde la columna 405 a la 425 las cadenas buscadas
	# si la encontramos usamos la funcion "replace" para reemplazar los caracteres encontrados por otros.
	# si no la encontramos, escribimos la linea tal cual la hemos leido.
	# la funcion "find" devuelve la posicion de inicio de la cadena buscada. Si no encuentra, devuelve -1.
	if cabecera == '54': 
		
		if cad2.find("17530102") > 0: 
			mod2 = cad2.replace("17530102", "19530101")
			fwrite.write(line[:405] + mod2 + line[425:])
			continue
		elif cad2.find("09741204") > 0:
			mod2 = cad2.replace("09741204", "19530101")
			fwrite.write(line[:405] + mod2 + line[425:])
			continue
		elif cad2.find("18610505") > 0:
			mod2 = cad2.replace("18610505", "19530101")
			fwrite.write(line[:405] + mod2 + line[425:])
			continue
		else:
			fwrite.write(line)
			continue

			
	# si la cabecera es 53 buscamos en los caracteres desde la columna 802 a la 822 las cadenas buscadas
	# si la encontramos usamos la funcion "replace" para reemplazar los caracteres encontrados por otros.
	# si no la encontramos, escribimos la linea tal cual la hemos leido.
	# la funcion "find" devuelve la posicion de inicio de la cadena buscada. Si no encuentra, devuelve -1.
	
	elif cabecera == '53':
	
		if cad1.find("17530102") > 0: 
			mod1 = cad1.replace("17530102", "19530101")
			fwrite.write(line[:802] + mod1 + line[822:])
			continue
		elif cad1.find("09741204") > 0:
			mod1 = cad1.replace("09741204", "19530101")
			fwrite.write(line[:802] + mod1 + line[822:])
			continue
		elif cad1.find("18610505") > 0:
			mod1 = cad1.replace("18610505", "19530101")
			fwrite.write(line[:802] + mod1 + line[822:])
			continue
		else:
			fwrite.write(line)
			continue
	else:
		fwrite.write(line)

fread.close()
fwrite.close()

	