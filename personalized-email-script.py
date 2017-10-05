import openpyxl
import textwrap

wb = openpyxl.load_workbook('Match Sheet.xlsx')

wb.get_sheet_names()

sheet = wb.get_sheet_by_name('Sheet1')

mentee = sheet['A3'].value.replace(" *","").replace("*","")
mentor = sheet['D3'].value.replace(" *","").replace("*","")
mentor_email = sheet['E3'].value
mentee_email = sheet['B3'].value

print textwrap.dedent("""\
Hi %s and %s,

Congratulations!

-About your mentor-
%s: %s
Rachel is a 3rd year biomedical mechanical engineering student. She loves art, hiking, running, sleeping, adventuring! She believes that the  advancements in self driving cars is something that will transform our future. An interesting thing on her bucket list is that she wants to go swimming with sharks even though she can't swim and is scared of sharks. How daring! She also loves R&B.


-About your mentee-
%s: %s
Reese is a 1st year biomedical mechanical engineering student. She love swimming, jogging, and badminton! She believes that 3D printing will transform our future. An interesting thing on her bucket list is that she wants to do CN Tower Walk. Sounds like fun! She also loves Rap/Hip Hop.

See you two soon!

-----------------------------------
On another note,

Are you interested in Power Electronics? IEEE Ottawa Section is hosting a Solantro Lab Tour for you! See the flyer attached. It's happening on October 10th!
      
      """) % (mentee.split()[0],mentor.split()[0], mentor, mentor_email, mentee, mentee_email)

from xlrd import open_workbook
book = open_workbook("Match Sheet.xlsx")
for sheet in book.sheets():
    for rowidx in range(sheet.nrows):
        row = sheet.row(rowidx)
        for colidx, cell in enumerate(row):
            if cell.value == "Sandy" :
                print "------------"
                print sheet.name
                print colidx
                print rowidx