import treepoem

data='OB001891312'
for each in ['azteccode', 'pdf417', 'code128', 'code39']:
    image = treepoem.generate_barcode(
        barcode_type=each,  # One of the BWIPP supported codes.
        data=data,
    )
    image.convert('1').save(each+data+'.png')
