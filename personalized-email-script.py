import openpyxl
import textwrap

wb = openpyxl.load_workbook('Match Sheet.xlsx')

wb.get_sheet_names()

sheet = wb.get_sheet_by_name('Sheet1')

mentee = sheet['A3'].value.replace(" *","").replace("*","")
mentor = sheet['D3'].value.replace(" *","").replace("*","")
mentor_email = sheet['E3'].value
mentee_email = sheet['B3'].value
mentor_year= "3"
mentee_year= "3"

she= True
if she ==True
    pronoun = "She"
else:
    pronoun = "He"


print textwrap.dedent("""\
Hi %s and %s,

Congratulations!

-About your mentor-
%s: %s
%s is a %s %s student. %s loves %s! %s believes that the  advancements in %s is something that will transform our future. An interesting thing on %s bucket list is that %s wants to %s. How exciting! %s also loves %s.

      """) % (mentee.split()[0],mentor.split()[0], mentor, mentor_email, mentor.split()[0], year, program)

print textwrap.dedent("""\      
-About your mentee-
%s: %s
%s is a %s year biomedical mechanical engineering student. She love swimming, jogging, and badminton! She believes that 3D printing will transform our future. An interesting thing on her bucket list is that she wants to do CN Tower Walk. Sounds like fun! She also loves Rap/Hip Hop.

See you two soon!

-----------------------------------
On another note,

Are you interested in Power Electronics? IEEE Ottawa Section is hosting a Solantro Lab Tour for you! See the flyer attached. It's happening on October 10th!
      
      """) % mentee, mentee_email)

# from xlrd import open_workbook
# book = open_workbook("Match Sheet.xlsx")
# for sheet in book.sheets():
#     for rowidx in range(sheet.nrows):
#         row = sheet.row(rowidx)
#         for colidx, cell in enumerate(row):
#             if cell.value == "Sandy" :
#                 print "------------"
#                 print sheet.name
#                 print colidx
#                 print rowidx