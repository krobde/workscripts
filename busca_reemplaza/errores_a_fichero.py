#import fileinput
#import pdb; 
#pdb.set_trace()

fread = open('RUSTICA_2018.txt', 'r')
fwrite = open ('errores_fechas2.txt', 'w')

for line in fread:
	cabecera = line[:2]
	cad1 = line[802:822]
	cad2 = line[405:425]
	
	if cabecera == '54':
	
		if cad2.find("17530102") > 0: 
			fwrite.write(line)
			continue
		elif cad2.find("09741204") > 0:
			fwrite.write(line)
			continue
		elif cad2.find("18610505") > 0:
			fwrite.write(line)
			continue
		else:
			continue

	elif cabecera == '53':
	
		if cad1.find("17530102") > 0: 
			fwrite.write(line)
			continue
		elif cad1.find("09741204") > 0:
			fwrite.write(line)
			continue
		elif cad1.find("18610505") > 0:
			fwrite.write(line)
			continue
		else:
			continue

fread.close()
fwrite.close()

	