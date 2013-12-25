import win32com.client
import os, sys


def findFlashFile(fileName):
	wpp = win32com.client.Dispatch('wpp.application')
	wpp.visible = True

	if wpp.Presentations.count > 1:
		print (wpp.Presentations.count)
		for i in range(wpp.Presentations.count):
			wpp.Presentations[0].close
			
	isHasFlash = False		
	pres = wpp.Presentations.Open(fileName, True)

	for slide_i in range(1, pres.slides.count + 1):
		slide = pres.slides(slide_i)
		for shape_i in range(1, slide.shapes.count + 1):
			shape = pres.slides(slide_i).shapes(shape_i)
			if shape.Type == 12:
				print ("slid{0}, shape{1} has flash".format(slide_i, shape_i))
				pres.close
				return True
	pres.close	
	return isHasFlash
	
def traveFile(fileDir):
	if os.path.exists(fileDir) == False:
		print ("path Error")
		return
		
	if os.path.isdir(fileDir) == True:
		list = os.listdir(fileDir)
		print (list)
		for item in list:
			item = os.path.join(fileDir, item)
			if os.path.isdir(item) == True:
				traveFile(item)
			else:
				index = item.rfind('.')
				suffix = item[index + 1:]
				print (suffix)
				if suffix.upper() in ["PPT", "PPTX", 'PPS', "DPS"]:
					hasFlash = findFlashFile(item)
					if hasFlash == True:
						file = open (sys.argv[2], "a")	
						file.write(item)
						file.write("\n")
						file.close

				else:
					print (item)
	
	
if __name__ == '__main__':
	print (sys.argv[1])
	traveFile(sys.argv[1])

		
	
	