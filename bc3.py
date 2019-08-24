from pubcode import Code128
barcode= Code128('OB00189137', charset='A')
barcode.image(height=40, module_width=1).save('pubcode-A.png')

barcode= Code128('OB00189137', charset='B')
barcode.image(height=40, module_width=1).save('pubcode-B.png')

barcode= Code128('OB00189137', charset='A')
barcode.image(height=40, module_width=1).save('pubcode-C.png')
