#  Email to PDF Conversion Macro

## Introduction
This is modified code taken from https://hinchley.net/ for converting emails and their attachments to one PDF document from within Outlook.

There are two core components to the solution:

- A module named Export. This code is responsible for exporting and merging emails and attachments. 
- A user form named Progress. This code provides a standalone progress dialog that is utilized by the Export module.

The operation of the macro has been created to perform the following tasks:

- Checks that a user has selected one or more emails. The macro exits with a warning message if no email has been selected.
- Prompts the user to set the path to where the generated PDF is to be saved. The macro will exit if no path is selected.
- Iterates through each selected email, saving the email, and each of the email’s attachments to the file system. Each email is saved as a Microsoft Word document, while each attachment is saved in its native format.
- If an attachment is an Excel workbook, each of the worksheets within the workbook are copied and pasted into a new Word document, which is then saved as a PDF.
- If an attachment is an Outlook message, the attachment is processed recursively (i.e. attachments within the attached email are also processed).
- The code then iterates through each of the saved documents, converting each to PDF format using Word. Note: If the document is an image, the image is pasted into a new Word document as an inline image, dynamically resized to avoid border-clipping, and then saved as a PDF.
- Finally, each of the PDF documents are merged into a single consolidated PDF using Adobe Acrobat DC.
- Each document is deleted after it is processed.

The following file types are supported:

- Word documents: doc, docx, docm, dot, dotm, dotx
- Excel documents: xls, xlsx, xlsm, xlt, xltm, xltx
- Images: jpg, jpeg, png, gif, bmp, tif, tiff
- Outlook emails: msg
- Other: txt, pdf, rtf

## Setup and Configuration
The following tasks need to be completed in Outlook for this to be setup.

#### Enable the Developer Tab
In Outlook go to File > Options > Customize Ribbon and tick the develper option under main tabs

#### Setup References
Open visual basic in Outlook and go to Tools > References and enable the following:

- Microsoft Word XX.0 Object Library 
- Microsoft Excel XX.0 Object Library
- Microsoft Scripting Runtime
- Microsoft VBScript Regular Expressions 5.5
- Microsoft Forms 2.0 Object Library
- Adobe Acrobat 10.0 Type Library
- Microsoft ActiveX Data Objects (multi-dimensional) 2.8
- OLE Automation

Depending on the version of office you have installed will depend on what version numbers the MS Office references contain.

| Version Name  | Version Number   |
| ------------ | ------------ |
| Outlook 2010  | 14.0   |
|  Outlook 2013 | 15.0  |
| Outlook 2016  | 16.0  |

#### Create Export & Progress VB
Create the following:

| Type  | Name  |
| ------------ | ------------ |
| Module  | Export  |
| UserForm  |  Progress |

When these have been created select "View Code" and copy the Export.vb and Progress.vb to their respective places

Note: Ensure that "ShowModal" is set to False for the Progress Userform

#### Apply a Digital Signature to Macro
For this to run while keeping the current security settings in place the macro will need to have a digital signature.
Complete the following steps to do this.
1. Go to C:\Program Files (x86)\Microsoft Office\root\Office16 and open SELFCERT.EXE
2. Name the certificate and press OK
3. Click OK to the notification stating the certificate was successfully created

Next, Open a Microsoft Management Console (MMC) and add the certificates snap-in with "my user account" selected.

Copy the certificate that has been created from: 

Certificates - Current User > Personal > Certificates

To:

Certificates – Current User > Trusted Root Certification Authorities > Certificates

Once this has been done close MMC and add the certificate to the VBA project. Go to Developer > Visual Basic and then go to Tools > Digital Certificate and select Choose to use the signature that has been created.

If problems are experienced with adding the certificate try the steps in this article: 

[Error Verifying VBA Project Signature](https://stackoverflow.com/questions/30619881/microsoft-outlook-2013-error-verify-vba-project-signature "Error Verifying VBA Project Signature")

#### Trust Centre Settings
The Macro settings in Trust Centre should be set to:

- Notifications for digitally signed macros. All other macros are disabled.

#### Quick Run Button
Setup a quick run button to run the Macro by doing the following:
- Go to File > Options > Quick Access Toolbar and select Macros from the choose commands dropdown
- Highlight the macro that has been created and press Add to move it over to the toolbar column.
- Click Modify to change the icon that is shown in the toolbar.

## How to Use
To use this select the email in Outlook you wish to convert and then click the quick run button you created.
Next select the place you wish to save the file and name it. 
Once this is done you will see the progress bar displayed and it will let you know when the email and it's attachements have been converted and saved.
Go to the location you saved the converted file and open it to make sure everythingis there.

## prerequisites
You will need the following installed:
- Microsoft Outlook
- Microsoft Word
- Microsoft Excel
- Either Acrobat Pro or Standard. (this will not work only with reader as you are creating a PDF)


