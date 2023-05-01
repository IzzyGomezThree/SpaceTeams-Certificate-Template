# SpaceCraft Certificate Generator

This code is used for two things: creating certificates and sending the respective certificates via email. Both are reliant on an Excel sheet.

## Description

For now, note that this code serves more as a template than an actual program. Feel free to copy the code and modify for your use. Code that does not have "template" in its name serves more as an example than a template, as the corresponding Excel files have been deleted to prevent personal information being public.

## Getting Started

### Languages Used
* Python3 (~95% of code)
* HTML (~5% of code)

### Dependencies

* Windows 10 (I think any operating system should work, such as Windows 11 or macOS, but I used Windows 10.)
* Python3
* A Python focused IDE (I used VS-Code, but other IDE's such as PyCharm should also work.)
* Microsoft Excel (MUST BE ABLE TO CREATE A .xlsx FILE.)

### Libraries

* [Pillow](https://pillow.readthedocs.io/en/stable/)
  - This library contains the functions used to edit images and save them. This is how I added text/lines to the certificates.
* [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/)
  - This library is what I used to read/write to an .xlsx Excel file.
* [redmail](https://pypi.org/project/redmail/)
  - This library contains functions used for automatically sending emails via Gmail. Outlook support is also supported.
* [pathlib](https://docs.python.org/3/library/pathlib.html)
  - Library used to edit file names and their location. I used this briefly for getting a file path.


## Explanations
1. Excel Organization
   - Only use a .xlxs file!!!! 
   - For the entire process, make sure to have columns for the following data:
     - First Name
     - Last Name
     - School Name
     - Certificate File Location
     - Voucher File Location
     - Email
     - (Any other relevant data)
   - If you are making certificates for teachers/mentors, make a separate worksheet and enter the relevant information there. 
   - Feel free to format the excel file as you see hit (color, column/row size, etc.)
2. File Formatting
   - The way I organized my files can be seen below:
   - ![FolderHierarchy](https://user-images.githubusercontent.com/47159819/220978437-ec1e80be-dd58-4a12-867f-444b9123fbac.jpg)
   - **Certificates** are where you place the generated certificates. If you are creating more than one type of certificates, create sub-folders within for all the types of certificates you generate.
   - **CertificateTemplates** are where the initial background images for the certificates are placed.
   - **EmailMedia** are where media used for the emails (photos or videos) are placed. 
   - **Fonts** are where all fonts used are stored.
   - **RealFiles** are where I placed my excel files and other files that I used to organize data.
   - All python files are created outside of the folders.
   - Feel free to reorganize things so that they make better sense to you, this is just a suggestion based on my experience.
3.  Certificate Workflow
    - Most of the workflow is a combination of trial and error, and a bit of planning. You plan the background template you want to use, the fonts that you want (including color, size, etc), and where/if you'll place the line. That being said, as you progress, you might change some of these factors, and that's where the trial and error comes in. 
4. Email Workflow
   - There are two parts to the email workflow:
     - **Email Draft:** Using an .html file, make sure to create a experiment on how the email will be formatted and how it will sound. Remember, these are professional emails going to students and mentors/teachers.  
     - **Sending Emails:** ALWAYS TEST BEFORE SENDING THE EMAILS!!! Before you officially send the emails, create a trial excel sheet and manually add the things you need for an email (email, file address, etc). For the trial, use your own email(s) and other Space Teams dev's. Once you and everyone think the email looks good, then call back the original email list and send the emails. DOUBLE CHECK via the senders email address (probably the spacecraft.vr gmail account) that the emails send; some emails may not have sent due to school's spam settings.
5.  Spacecraft Email Password
    - For the email python file, you will be creating an App Password, which is basically a temporary generated password that can be used to access the Spacecraft gmail account. Because of this, make sure the password is stored in a safe place, and that no one besides a Spacecraft Lead has access to the email python file (unless given permission). For access to the email, talk to Dr.Chamitoff.

## Step-By-Step: Prerequisites

Note: most these steps can be skipped if downloading from this GitHub,with the exception of steps 1 & 8

1. Make sure all Dependencies (above) are downloaded, including the latest version of Python3, an IDE for python, and Microsoft Excel. 
2. Create a new folder somewhere accessible on your computer called "SpaceTeamsCertificates[Semester####]"; for example, if it's in Summer 2023, you would name it "SpaceTeamsCertificatesSummer2023".
3. Under that folder, make 5 different folders called "Certificates", "Certificate Templates", "EmailMedia", "Fonts", and "RealFiles".
4. Under "RealFiles", create a .xlxs file called "SpaceTeams[Semester####]Info"; for example, if it's in Summer 2023, you would name it "SpaceTeamsSummer2023Info.xlxs".
   - Respectfully name individual columns "FirstName", "LastName", "SchoolName", "CertificateFileLocation", "VoucherFileLocation", and "studentEmail". 
   - If available already, fill out the excel sheet with the respective information of each student/participant; each student should have their own row. Make sure to save with Ctrl + S! (If you do not have the student's information yet, make sure to get that information and fill out the excel sheet asap!)
   - If you are also making certificates for teachers/mentors, create a new worksheet, and copy the respective columns. 
   - A template file for this can be found under "RealFiles -> Template-Excel-File.xlxs".
5. Download the "Freehand-Regular.ttf", "arial-bold.ttf", "arial.ttf", and "open-sans.semibold.ttf" fonts from this GitHub, and place them in the "Fonts" folder. 
6. Download "blankCertificate_v2.png" and "studentCertificate_v2.png" from this GitHub, and place them in the "CertificateTemplates" folder.
7. Download the "spaceTeams_email.png" from this GitHub, and place it in your "EmailMedia" folder.
8. Setup your respective python IDE by making sure it can run python scripts and installing the Libraries (listed above).
   - For instructions on installing the libraries, look through the links above under Libraries. 

## Step-By-Step: Certificate Creation

Note: This is a step-by-step process for creating certificates. If you haven't already, make sure to complete the "Step-By-Step: Prerequisites" before starting this section. 

1. Open "certifcateCreation_Template.py"; read the top for more information about how files work.
2. Install Pillow and OpenPyxl libraries (if you haven't already).
3. Declare the libraries in the python file and their respective imports.
4. Access the excel workbook (line 24), access the first worksheet (line 25 & 26), and declare the max row (line 27). 
5. Start a For Loop, ranging from the first row (1) to the last row (n_row + 1).
6. Access the raw image.
7. Declare the photo size.
8. Call draw Method to add 2D graphics in an image.
9. Read values from excel sheet (as tuple).
10. Turn tuple variable to strings, and combine first and last name into one string. 
11. Decalre font and font size.
12. Declare dynamic font variables. 
13. Write code that decreases the font size based on the img_fraction variable.
14. Declare line that goes under the name of participant. 
15. Declare code that increases the line size for larger names.
16. Add text and line to the photo.
17. Save image into the "Certificates" folder.
18. Add the image file path to the excel sheet in column D.
19. Add debug code that shows in command line which certificate is completed.

## Step-By-Step: Emailing Certificates

Note: This is a step-by-step process for emailing your generated certificates. If you haven't already, make sure to complete the "Step-By-Step: Prerequisites" and "Step-By-Step: Certificate Creation" before starting this section. 

1. Open "emailCreation_Template.py"; read the top for more information about how files work.
2. Install redmail, pathlib, OpenPyxl libraries (if you haven't already).
3. Declare the libraries in the python file and their respective imports.
4. Access the excel workbook (line 22), access the first worksheet (line 23 & 24), and declare the max row (line 25).
   - KEEP LINE 22 COMMENTED! THIS IS TO PREVENT ACCIDENTLY SENDING ALL EMAILS/CERTIFICATES PREMATURELY! IF YOU WANT TO TEST YOUR CODE, READ "Email Workflow" UNDER "Explanations" ABOVE FOR GUIDANCE!!!
5. Start a For Loop, ranging from the first row (1) to the last row (n_row + 1). 
6. Read values from excel sheet (as tuple).
7. Turn tuple variable to strings. 
8. If an excel cell is empty, continue to the next row.
9. Get rid of all file path clutter before the certificate PNG file.
   - This is done so that way when the participants get the certificate via an email attachment, they see "IsraelGomezCertificate.png", as opposed to "C:/Users/..../IsraelGomezCertificate.png".
10. Log into the senders (aka you) Gmail account.
    - There is a debugging email section commented out that can be used for testing email log in.
    - If using a debugging email, comment out the Spacecraft email.
11. Start the email process
12. Include subject line, sender, receiver, and attachment information.
13. Include an email in HTML form.
    - For information on creating the HTML email, look at "How do you create a HTML email?" below under "Common Issues/Questions - Emailing Certificates".
14. Use debugger to see when an email has been sent.




## Common Issues/Questions - Certificate Creation

1. Where do I get a certificate template photo?
   - Assuming the template stays the same, keep using the "blankCertificate_v2.png" under "CertificateTemplates".
2. How do I get the template photo size?
   - Use the following code below (shamelessly stolen from [here](https://stackoverflow.com/questions/6444548/how-do-i-get-the-picture-size-with-pil)), and note down the values.
```python
from PIL import Image

im = Image.open('whatever.png')
width, height = im.size
```
3. Where do I get fonts?
   - A quick Google search should do the trick. For example, if you're looking for the Ariel font, look up "arial font download", and find a download; I usually used [Fonts Family](https://freefontsfamily.com/). Note that getting **bold** or *italic* variations require another/separate download.
4. How do I choose font size?
   - Experimentation.
5. How does the dynamic text work?
   - This function is used for participant's/school's with long names. Basically, I first started off by declaring the original font size, and the percentage size of the text; so if I choose 95%, the text covers 95% of the width of the certificate. 
```python
# declare dynamic font size variable
img_fraction = 0.80
fontsize = 100
    
# when the text goes out of bounds from the photo size, it slowly decreases until the text fills 95% of the photo)
freehandFont = ImageFont.truetype("Fonts/Freehand-Regular.ttf", fontsize)
while freehandFont.getsize(fullName_value)[0] > img_fraction*img.size[0]:
     # de-increment to be sure it is less than criteria
      fontsize -= 1
      freehandFont = ImageFont.truetype("Fonts/Freehand-Regular.ttf", fontsize)
```
6. How does the dynamic line work?
   - This line is what is put under a name for formality; there is logic to prevent the line from going ugly across the certificate, but for it to also be slightly bigger than the name. Basically, the line is declared, and will increase in size based on how long the name is, and will always be 60 pixels longer than the name; however, it will never be larger 1064 pixels, to avoid going across the entire certificate. Note that these numbers will have to be changed if the certificate resolution changes. 
```python
# declare line; (0,0) is top left corner of the certificate
y = 415     # the higher the value, the lower the line
delta_x = 400   # the higher the value, the wider the line
x1 = 560.5 - delta_x   # x-coordinate on the left side
x2 = 560.5 + delta_x   # x-coordinate on the right side
width_x  = x2-x1 
shape = [(x1, y), (x2, y)]    # 560.5 is the midway point on the x-axix
    
# dynamically increase the line size when the text surpasses the border by 60 pixels
while freehandFont.getsize(fullName_value)[0] > (width_x-60):
      if width_x < 1064:      # 1064 pixels is the longest width the line is allowed to be
         delta_x += 1
         x1 = 560.5 - delta_x   # x-coordinate on the left side
         x2 = 560.5 + delta_x   # x-coordinate on the right side
         width_x  = x2-x1 
         shape = [(x1, y), (x2, y)]    # 560.5 is the midway point on the x-axix
      else:
         break
```
## Common Issues/Questions - Emailing Certificates

1. How do you create a HTML email?
   - First off, make a .html file. Then, using a IDE (such as VS-Code) or this [online editor](https://www.tutorialspoint.com/online_html_editor.php), create your HTML email. Once that is done, you have two ways (that I can think of) to make an email.
     * **Option 1:** Manually create an email using HTML.
     * **Option 2:** Draft a email using Gmail. Once you are happy with your email, highlight the entire email, right-click it, and press "Inspect". Copy the entire highlighted HTML code into your HTML file. 
2. Why won't the email(s) send?
   - More than likely, it's because your Google account App Password isn't generated or is old. Make sure to follow [these steps](https://red-mail.readthedocs.io/en/stable/tutorials/config.html#config-gmail) to (re)generate a new App Password.

## Authors

* Israel Gomez, izzygomezthree@tamu.edu
* Neil McHenry

Feel free to message Israel for any questions, comments, or recommendations. 

## Future Plans

* Creating an actual executable for both the certificate generation and email automation.
* Creating a more artistic and eye-pleasing email template.

## Version History

* 0.1
    * Initial Release

## License

[MIT](https://choosealicense.com/licenses/mit/)

## Acknowledgments

* [SpaceCRAFT](https://spacecraft-vr.com/)
* [Astro Center](https://astrocenter.tamu.edu/)
* [Texas A&M University](https://engineering.tamu.edu/aerospace/index.html)
* [DomPizzie](https://gist.github.com/DomPizzie/7a5ff55ffa9081f2de27c315f5018afc) for README template
* [Make a README](https://www.makeareadme.com/) for helping me draft my README file
* Israel's cat Daisy, and Neil's dog Scout

![received_3034991500139329](https://user-images.githubusercontent.com/47159819/221294613-ee59f482-43d5-46ee-a2a9-00b744a0d1bd.jpeg) ![20210911_151523](https://user-images.githubusercontent.com/47159819/221294634-2b020ca0-f8af-4431-baf1-3b5ef539094c.jpg)
