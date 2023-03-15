import os
import pandas as pd

workbook = pd.read_excel('TEMP-ITG_AllOrganizations.xlsx', usecols = 'A:G')
workbook.head()

index = 0
number = 240

TestArray = [['Name', 'Domains', 'SPF'], ['', '', '']]


with open("output.txt", 'w') as file:
    while number > 0:
        try:
            #print(index)
            file.write("\n")
            file.write(str(workbook['name'].iloc[index]))
            print(str(workbook['name'].iloc[index]))
            file.write("\n")
            if str(workbook['Domain'].iloc[index]) == 'nan': #or str(workbook['Domain'].iloc[index]) == None or str(workbook['Domain'].iloc[index]) == '':
                file.write("[No Website Provided.]")
                print("[No Website Provided]")
            else:
                os.system(f"python spf.py {str(workbook['Domain'].iloc[index])}")
                
                with open('temp.txt', 'r') as file2:
                    spf = str(file2.read())
                file2.close
                flag =1
                
                #print(str(workbook['Domain'].iloc[index]))
                file.write(str(workbook['Domain'].iloc[index]))
                
                if str(workbook['Domain2'].iloc[index]) != 'nan':
                    #print(str(workbook['Domain2'].iloc[index]))
                    file.write(', '+str(workbook['Domain2'].iloc[index]))
                    flag =2
                if str(workbook['Domain3'].iloc[index]) != 'nan':
                    #print(str(workbook['Domain3'].iloc[index]))
                    file.write(', '+str(workbook['Domain3'].iloc[index]))
                    flag =3
                if str(workbook['Domain4'].iloc[index]) != 'nan':
                    #print(str(workbook['Domain4'].iloc[index]))
                    file.write(', '+str(workbook['Domain4'].iloc[index]))
                    flag =4
                
                file.write(f'\n{spf}')

                
                if flag == 1:
                    TestArray.insert(2+index, [str(workbook['name'].iloc[index]), str(workbook['Domain'].iloc[index]), spf])
                elif flag == 2:
                    TestArray.insert(2+index, [str(workbook['name'].iloc[index]), str(workbook['Domain'].iloc[index])+', '+str(workbook['Domain2'].iloc[index]), spf])
                elif flag == 3:
                    TestArray.insert(2+index, [str(workbook['name'].iloc[index]), str(workbook['Domain'].iloc[index])+', '+str(workbook['Domain2'].iloc[index])+', '+str(workbook['Domain3'].iloc[index]), spf])
                elif flag == 4:
                    TestArray.insert(2+index, [str(workbook['name'].iloc[index]), str(workbook['Domain'].iloc[index])+', '+str(workbook['Domain2'].iloc[index])+', '+str(workbook['Domain3'].iloc[index])+', '+str(workbook['Domain4'].iloc[index]), spf])
            file.write("\n")
            index +=1
            number -=1
        except:
            print("Done.")
            break
    
file.close()

for r in TestArray:
    for c in r:
        if c == None:
            c = ''

df = pd.DataFrame(TestArray)
writer = pd.ExcelWriter('output.xlsx')
df.to_excel(writer, sheet_name='TEMP-ITG_AllOrganizations_output', index=False)
writer.save()
