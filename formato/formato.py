path = '/mnt/d/borra/DOC_AGO_17.text'
file_orig = open(path,'r')
#days = file_orig.read()

new_path = '/mnt/d/borra/DOC_AGO_17.txt'
mod_file = open(new_path,'w+')

with open(path) as f:
	for line in f:

		#linea = f.readline()
		#vari = "{:<1000}".format(linea)
		vari = (line.strip()).ljust(1000)+'\n'
		
		# print '%1000s' % file_orig.readline(lineas)	 
		#z="{:<1000}\n".format(lineas)
		#mod_file.writelines("%-1000s" % lineas)
		
		#mod_file.write('{:<1001}'.format(line))
		#mod_file.write('{:<1000}'.format(f.readline()))
		mod_file.write(vari)

file_orig.close()
mod_file.close()