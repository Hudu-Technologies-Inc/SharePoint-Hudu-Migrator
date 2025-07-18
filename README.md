# Hudu Sharepoint Migration
Easy Migration from Sharepoint to Hudu

## Getting Started

### Prerequisites

- Hudu Instance of 2.37.1 or newer
- Companies created in Hudu if you want to attribute sharepoint items to companies
- Hudu API Key
- Sharepoint / Sites
- Read permissions for Sharepoint
- Powershell 7.5.0 or later
- Libreoffice Installed

### Starting

It's reccomended to instantiate via dot-sourcing, ie

```
. .\Sharepoint-Migration.ps1
```

it will check powershell version, get modules loaded, get you signed into hudu, check hudu version, and begin downloading all sites/files.

you'll be asked to copy a code for authenticating via Microsoft Device Login. Simply copy the generated code and paste in after navigating [here](https://login.microsoftonline.com/common/oauth2/deviceauth) in a web browser. You'll then sign into Office/Azure/Entra as usual.

Just before the file conversion process begins, this script will download a and being the graphical installer for Libreoffice if it doesn't seem like Libreoffice is installed. Simply go through all the questions with default values, as the desired install path has already been set for you. If you already have Libreoffice installed, it will pick up on that in use the version you already have.


### Setup Questions

#### Question 1 - **From Which Site(s) Would You Like to Get Source Materials From**
- transfer everything **from a single Sharepoint site**
- transfer everything **from some Sharepoint sites**
- transfer everything **from all Sharepoint sites**

#### Question 2 - **To Which Companies, if any Would You Like to Transfer To**
- transfer everything **to a single Hudu Company**
- transfer everything **to central / global KB (no company association)**
- transfer everything **to companies / central kb of my choosing (select a company / destination for each article)**

#### Question 3 - **Would you like to include links to original SharePoint Documents**
- Yes, **include a link in each Hudu article to original files** in Sharepoint
- No, **we don't need links in Hudu to original Sharepoint Files**

#### Question 4 - **Would you like to include a copy of original Sharepoint Document as Attachments to Hudu Articles**
- **Upload original files** from sharepoint to Hudu Articles (for posterity/reference)
- No need, **don't upload original documents** to new Articles
<img width="392" height="33" alt="image" src="https://github.com/user-attachments/assets/5a8460a4-e7c7-4d03-98fd-ba2dd935bf4b" />

#### Question 5 - **Would you like to Convert Spreadsheets to Articles?**
- **Convert spreadsheets to articles**, effectively making an html table for this data
- **Don't convert spreadsheets**, just attach/upload them to Hudu 
<img width="359" height="28" alt="image" src="https://github.com/user-attachments/assets/8e6bf587-6027-4a2b-846b-39be6710d4e0" />

#### Question 6 - **Would you like to Convert PowerPoints to Articles?**
- **Convert powerpoints to articles**, effectively making an html table for this data
- **Don't convert powerpoints**, just attach/upload them to Hudu 
<img width="359" height="28" alt="image" src="https://github.com/user-attachments/assets/8e6bf587-6027-4a2b-846b-39be6710d4e0" />


## Supported Files?

All files can be added if they are either
-under 100mb in size
-under 8500 Characters long when converted to html

Otherwise we just link to original file on Sharepoint

### Document files
#### work nicely and any of the below should convert nicely:
```
- ".pdf"
- ".doc"
- ".docx"
- ".docm"
- ".rtf"
- ".txt"
- ".md"
- ".wpd"
- ".odt"
```

### Spreadsheets and Tabular Data
#### can be converted if you like, but any formulae will be replaced. If you plan on editing tabular data frequently, you might elect to not convert these"
```
- ".xls"
- ".xlsx"
- ".csv"
- ".ods"
```

### Powerpoints and Presentations
#### These can be converted to a simplified html equivilent. Each slide has it's data / images extracted
```
- ".ppt"
- ".pptx"
- ".pptm"
- ".odp"
```

### All others:
#### All others will skip conversion process and will be uploaded directly if **under 100mb in size**. If they are **over 100mb in size**, we just create a stub article that has a link to the sharepoint file in question.

All other non-document files skip conversion and are uploaded with a simple article containing file details.
<img width="723" height="45" alt="image" src="https://github.com/user-attachments/assets/619859ee-b367-483d-85ed-b2336f7eda34" />
most common formats that we skip conversion for are:
```
".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a", ".dll", ".so", ".lib", ".bin", ".class", ".pyc", ".pyo", ".o", ".obj", ".exe", ".msi", ".bat", ".cmd", ".sh", ".jar", ".app", ".apk", ".dmg", ".iso", ".img", ".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".tgz", ".lz", ".mp4", ".avi", ".mov", ".wmv", ".mkv", ".webm", ".flv", ".psd", ".ai", ".eps", ".indd", ".sketch", ".fig", ".xd", ".blend", ".ds_store", ".thumbs", ".lnk", ".heic"
```

Files that **don't have an extension**, we'll attempt to decode these as UTF-8

Files that aren't traditional documents will have a page generated for them with your chosen links. here's what an image in sharepoint will look like after adding to hudu 
<img width="1184" height="1078" alt="image" src="https://github.com/user-attachments/assets/6a8fc563-49ee-4ed6-b17a-b35c78cc42f1" />

### If a file is too large after conversion (longer than 8500 characters)
if this condition is met, the original is uploaded in Hudu and Linked
 <img width="849" height="363" alt="image" src="https://github.com/user-attachments/assets/4354122b-2f97-41ce-b64e-9e0c6262072e" />

#### Effectively, here's the process:
Select desired Sharepoint Site(s)
Download all files from selected sites
convert those that can be converted
Those that can be converted- Parse links, extract images, etc
Those that can't - generate a useful landing page for file


## FAQs:

#### Q: Does my original Sharepoint Folder Structure Carry Over?
A: Yes. If you select a company to attribute an article to, the original folder structure is retained for that article
<img width="822" height="449" alt="image" src="https://github.com/user-attachments/assets/4e34671d-e769-4553-9566-e573ffb03720" />

#### Q: What does a PowerPoint file look like after converting into Hudu?
A: Pretty basic. Information from each slide is extracted, including images. This information is extracted into a web-friendly Article
<img width="896" height="498" alt="image" src="https://github.com/user-attachments/assets/3ba89439-61fb-4995-a46e-2fd2595b6ab7" />

#### Q: What does a Spreadsheet file look like after converting into Hudu?
A: It constructs an HTML table from your spreadsheet. Large spreadsheets are less manageable and any-sized spreadsheet will lose its formulae. It can be good for making sure information from Sharepoint is searchable and findable in Hudu, however.
<img width="901" height="483" alt="image" src="https://github.com/user-attachments/assets/1d3188bc-495a-4553-8d81-ade5a349ee4c" />



