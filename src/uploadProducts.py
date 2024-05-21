from functions import signInWooAB,getFolders,downloadOneProduct,NewProduct
from time import sleep
wcapi = signInWooAB()
pathWebImages=wcapi.url+'wp-content/uploads/'
talla=3     # ID variation for talla: 3 normal, 1 bebe
colour=5    # ID variation for colour

'''We get all data from the .txt files
    everything should be in the products2upload'''
folder='test'
prod_folders=getFolders(folder)
print(prod_folders)


for prodFolder in prod_folders:
    prodFolder2=prodFolder.split('\\')
    prodFolder=prodFolder2[-2]
    print(prodFolder)

    NewProductEcommerce = NewProduct()
    NewProductEcommerce.readFromCsv(folder, folder1=prodFolder)
    # Now we create the json to upload all the PRODUCT information
    NewProductEcommerce.upload_product(wcapi ,pathWebImages, talla, colour, toUpload=1, imagesProduct = False)

    sleep(5)

