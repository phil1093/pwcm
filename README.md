#pwcm
import openpyxl #import modules
import os

#set up and open excel file needed, assigning sheet names

os.chdir('/Users/philipwhitfield/Google Drive/Panini')
wb = openpyxl.load_workbook('wcm.xlsx')
sheet1 = wb.get_sheet_by_name('Duplicates')
sheet2 = wb.get_sheet_by_name('Need')


#W35 is the furthest cell in sheet1 Duplicates
#Collumn 1 is Countries
#column 2 is number on sticker
#35 rows total,
#4th column starts country at 0,22,32,52,72,etc..
#23rd column is final column (W) ending 21,31,52,71 etc..

#Stadium values
#for i in range(1,2):
#	for j in range(1,24):
#		print(i, j, sheet1.cell(row=i, column=j).value)

#get card number


#loop over for number of cards
while True:
    cardnum = int(input( 'give card number '))
    if cardnum == "":
        break
    #Create column number
    coln = [int(n) for n in str(cardnum)]

    cardnumstr = [int(n) for n in str(cardnum)]

    if cardnum < 100:
        if cardnumstr[0] % 2 ==0:
            column = coln[1] + 12
        elif cardnumstr[1] in range(2):
            column = coln[1] + 22
        else:
            column = coln[1] + 2
            
    else:
        if cardnumstr[1] % 2 ==0:
            column = coln[2] + 12
        elif cardnumstr[2] in range(2):
            column = coln[2] + 22
        else:
            column = coln[2] + 2
    #print('column is ' + str(column))
    #generating row number
    row = 2

    if cardnum in range(1,31):
        row = row
    else:
        row = int(3 +( (cardnum - 32) / 20))    
    #print('row is ' + str(row))

    #convert column num to excel letter format
    def Col2xl(idx):
        if idx < 1:
            raise ValueError("Index is too small")
        result = ""
        while True:
            if idx > 26:
                idx, r = divmod(idx - 1, 26)
                result = chr(r + ord('A')) + result
            else:
                return chr(idx + ord('A') - 1) + result

    col = Col2xl(column)           
    celllocation = col + str(row)
    print(celllocation)

    #have location of new card number,
    #need to check if cell in sheet2 is empty or not
    #then if empty to write new card number in sheet1
    #Else if it is full, to delete the cell in sheet2
    #here goes..


    if sheet2[celllocation].value == None:
        sheet1[celllocation] = cardnum
        print('it is a duplicate')
    else:
        sheet2[celllocation] = None
        print('New Card')

    #save new file

    wb.save('wcm.xlsx')
