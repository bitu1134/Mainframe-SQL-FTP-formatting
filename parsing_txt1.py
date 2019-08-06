import glob,os,pandas as pd,time,easygui
from datetime import datetime
from xlwt import Workbook

#name=input("Desired output name ")
os.chdir(r"C:\Users\manoranjann\Documents\to_parse")
filename=r"C:\Users\manoranjann\Documents\to_parse\OUTPUT1.csv"
#tifCounter = len(glob.glob("*.txt"))
#print(tifCounter)
i=1
filenm=[]
a14 = easygui.ynbox('Do you want to generate output for the recently addded file', 'Confirmation', ('Yes', 'No'))

if a14==True:
    name = easygui.enterbox("Please enter desired file output name ")
    for file in glob.glob('*.txt'):
        print(datetime.strptime(time.ctime(os.path.getmtime(file)),"%a %b %d %H:%M:%S %Y"))
        filenm.append(datetime.strptime(time.ctime(os.path.getmtime(file)),"%a %b %d %H:%M:%S %Y"))
        i+=1
        print(file+' ' + str(datetime.strptime(time.ctime(os.path.getmtime(file)),"%a %b %d %H:%M:%S %Y")))
    print(max(filenm))
    mx=max(filenm)
    a=0
    for file in glob.glob('*.txt'):
    #  print(file)
      if datetime.strptime(time.ctime(os.path.getmtime(file)),"%a %b %d %H:%M:%S %Y")==mx:
        print('FILE WHICH IS PROCESSED for header- %s' %file)
        fread = open(file)
        for line in fread:
            if '---------+' in line:
               a+=1
               if a==2:
                   header=next(fread)
                   break

    b=0
    fwrite = open(filename, 'w')
    fwrite.writelines(header)
    for file in glob.glob('*.txt'):
      if datetime.strptime(time.ctime(os.path.getmtime(file)), "%a %b %d %H:%M:%S %Y") == mx:
        print('FILE WHICH IS PROCESSED for output file- %s' % file)
        mxfile=file
        fread = open(file)
        for line in fread:
            if header in line:
               b+=1
            if b>0 and '---------+' not in line and header not in line and line.startswith('DSNE6')== False:
                fwrite.writelines(line )
    fwrite.close()
    #
    try:
    # input("Press enter to continue")
    # easygui.msgbox("Click OK to  confirm", title="Info Message")
     a = easygui.ynbox('Do you want to generate output of file %s' %mxfile, 'Confirmation', ('Yes', 'No'))
    except SyntaxError:
        pass
    if a==True:
        a2 = easygui.ynbox('Is this message related table file - %s' % mxfile, 'Confirmation', ('Yes', 'No'))
        if a2==True:
            df = pd.read_fwf(r"C:\Users\manoranjann\Documents\to_parse\OUTPUT1.csv")
            df.to_csv(r"C:\Users\manoranjann\Documents\to_parse\%s.csv" % name, index=None, header=False,sep=' ')
            fread = open(r"C:\Users\manoranjann\Documents\to_parse\%s.csv" % name)
            wb = Workbook()
            sheet1 = wb.add_sheet("Sheet 1", cell_overwrite_ok=True)
            a = 0
            b = 0
            for line in fread:

                if '-}' in line:
                    print('coming')
                    sheet1.write(b, a, line)
                    a += 1
                    b = 0
                else:
                    # print(df.iloc[i,0])
                    sheet1.write(b, a, line)
                    b += 1
            # print(wb)
            wb.save(r'C:\Users\manoranjann\Documents\to_parse\%s.xls' %name)
            easygui.msgbox("file generated in  \Documents\ to_parse\ ", title="XLS GEN")
            a11 = easygui.ynbox('Do you want to open the file ', 'Confirmation', ('Yes', 'No'))
            if a11 == True:
                os.startfile(r"C:\Users\manoranjann\Documents\to_parse\%s.xls" % name)
        else:
            df = pd.read_fwf(r"C:\Users\manoranjann\Documents\to_parse\OUTPUT1.csv")
            df.to_csv(r"C:\Users\manoranjann\Documents\to_parse\%s.csv" %name, index = None, header=True)
            easygui.msgbox("file generated in  \Documents\ to_parse\ ", title="CSV GEN")
            print("file generated in  \Documents \ to_parse\ ")
            a1 = easygui.ynbox('Do you want to open the file ' , 'Confirmation', ('Yes', 'No'))
            if a1==True:
                os.startfile(r"C:\Users\manoranjann\Documents\to_parse\%s.csv" %name)
            time.sleep(2)
    else:
        easygui.msgbox("Process Interrupted", title="Info Message")
else:
    easygui.msgbox("Process Interrupted", title="Info Message")