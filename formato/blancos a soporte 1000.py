#Ejecutamos abriendo cmd en windows, ejecutamos bash y luego python mas nombre del programa.py
#Este programa formatea cada línea a 1000 caracteres, anadiendo espacios al final del fichero.

path = '/mnt/d/borra/JUN_17.txt.old' 	#definimos el fichero origen
file_orig = open(path,'r')				#abrimos el fichero origen
#days = file_orig.read()

new_path = '/mnt/d/borra/JUN_17_mod.txt' #definimos el fichero destino
mod_file = open(new_path,'w+')			 #abrimos el fichero destino. w+ crea el fichero si no existe


#recorremos el fichero.
with open(path) as f:
	for line in f:

		#linea = f.readline()
		#vari = "{:<1000}".format(linea)
		vari = (line.strip()).ljust(1000)+'\n'		#utilizamos strip porque el programa añadía un caracter en blanco, no se motivo.
													#ljust crea la línea con el ancho pasado como parámetro.
		
		# print '%1000s' % file_orig.readline(lineas)	 
		#z="{:<1000}\n".format(lineas)
		#mod_file.writelines("%-1000s" % lineas)
		
		#mod_file.write('{:<1001}'.format(line))
		#mod_file.write('{:<1000}'.format(f.readline()))
		mod_file.write(vari)

file_orig.close()
mod_file.close()