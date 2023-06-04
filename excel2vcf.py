import pandas,sys,os

if len(sys.argv) == 1:
   print("The excel file is not specified\nUsage: excel2vcf.exe contacts.xlsx") 
   sys.exit()
df = pandas.read_excel(sys.argv[1])


if os.path.exists("contacts.vcf"):
        os.remove("contacts.vcf")
        print("contacts.vcf file found, it will be replaced by the version")

for name,phone in zip(df[df.columns[0]].values,df[df.columns[1]].values ):
   phone = str(phone)
   if len(phone) == 9:
        phone='0'+phone
   
    
   f = open("contacts.vcf", "a")
   f.write("BEGIN:VCARD"+"\nVERSION:3.0"+"\nN:;"+name+";;;"+"\nFN:"+name+"\nTEL;type=Home:"+phone+"\nEND:VCARD\n")
   f.close() 
 
print("file contacts.vcf created successfully") 

