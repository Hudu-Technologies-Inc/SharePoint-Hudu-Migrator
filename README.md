# Hudu SharePoint Migration
Easy Migration from SharePoint to Hudu

## Note on Intended Purpose
If you are working from a single mapped SharePoint drive rather than multiple SharePoint sites, you may prefer the direct [Files-Hudu-Migration](https://github.com/Hudu-Technologies-Inc/Files-Hudu-Migration) tool. However, for environments with multiple SharePoint sites, this project is the recommended solution. It preserves inter-site links and treats each SharePoint site as a distinct entity, whereas Files-Hudu-Migration processes each file or folder as an individual entity.

### Prerequisites

- Hudu Instance of 2.37.1 or newer
- Companies created in Hudu if you want to attribute sharepoint items to companies
- Hudu API Key
- SharePoint / Sites with Files
- Powershell 7.5.0 or later
- Libreoffice **(script will start install if not present)**
- Read permissions for SharePoint **(script will register app in entra if not manually set)**

### Environment File and Invocation

> **Permissions Notice**
>
> Some scripts may require elevated permissions. If you encounter access-related errors, consider launching PowerShell (`pwsh`) with **Run as Administrator**.
>
> Please note that administrative privileges do not override Windows Rights Management or similarly enforced file protection mechanisms.

If you are using an environment file to hold your secrets and user-variables, you can make a copy of the environment file example and enter your values as needed.
```
Copy-Item environment.example my-environment.ps1
notepad my-environment.ps1
```

And then you can kick it off by opening `pwsh7` session as `administrator`, and `dot-sourcing` your `environment file`, as in the below example-

```
. .\my-environment.ps1
```

Alternatively, if you don't wish to fill out an environment file, you can invoke this script directly and you'll be asked for these values as they are needed.
Kick off this script directly by opening `pwsh7` session as `administrator`, and `dot-sourcing` the Sharepoint Migration Script

```
. .\Sharepoint-Migration.ps1
```


### Setup Azure AppRegistration

#### Auto-Registration
If you haven't entered your ClientId (AppId) and TenantId from Microsoft, you can complete Azure App Registration within this script. Please follow the prompts carefully for a smooth process.

<img width="1000" height="500" alt="image" src="https://github.com/user-attachments/assets/d3c384f8-9610-4977-8338-a102b7f8f87c" />

Since we are using Microsoft's `MSAL authentication module` with `Device Auth`, you'll need to ensure that 'Public Client Flows' is switched-on. You can do this manually, otherwise this blade will be opened up automatically.

**At the end, you'll be given the option to remove this app registration**

#### Manual-Registration
Otherwise, if you want to set up Azure/Entra AppRegistration manually, you'll need to set up with the permissions below, then place the ClientID (AppId) and TenantID in the top of the main script.
```
delegatedPermissions: "Sites.Read.All", "Files.Read.All", "User.Read", "offline_access"
applicationPermissions: "Files.Read.All", "Sites.Read.All"
```
<img width="500" height="100" alt="image" src="https://github.com/user-attachments/assets/0dccc77c-25b3-4a55-99f4-aa4ed5e8dcbb" />

## Getting Started
This script will check the PowerShell version, get all the necessary modules loaded, prompt you to sign into Hudu, check the Hudu version, and make sure you're authenticated to SharePoint.

You'll be asked to copy a code for authenticating via Microsoft Device Login. Simply copy the generated code and paste it after going to Microsoft's device authentication page, [here](https://login.microsoftonline.com/common/oauth2/deviceauth) in a web browser. Sign into Office/Azure/Entra as you usually would.

[<img width="1972" height="1184" alt="image" src="https://github.com/user-attachments/assets/02041a8d-b7ce-48f7-aa90-248d16798e3f" />](https://login.microsoftonline.com/common/oauth2/deviceauth)

Just before the file conversion process begins, this script will download and install the graphical installer for Libreoffice, if not install already. Simply go through all the questions with default values, as the desired install path has already been set for you. If you already have Libreoffice installed, it will pick up on that and use the version you already have.


### Setup Questions

#### Question 1 - **From Which Site(s) Would You Like to Get Source Materials From**
- transfer everything **from a single SharePoint site**
- transfer everything **from some SharePoint sites**
- transfer everything **from all SharePoint sites**

#### Question 2 - **To Which Companies, if any Would You Like to Transfer To**
- transfer everything **to a single Hudu Company**
- transfer everything **to central / global KB (no company association)**
- transfer everything **to companies / central kb of my choosing (select a company / destination for each article)**

#### Question 3 - **Would you like to include links to original SharePoint Documents**
- Yes, **include a link in each Hudu article to original files** in SharePoint
- No, **we don't need links in Hudu to original SharePoint Files**

#### Question 4 - **Would you like to include a copy of original SharePoint Document as Attachments to Hudu Articles**
- **Upload original files** from sharepoint to Hudu Articles (for posterity/reference)
- No need, **don't upload original documents** to new Articles

#### Question 5 - **Would you like to Convert Spreadsheets to Articles?**
- **Convert spreadsheets to articles**, effectively making an html table for this data
- **Don't convert spreadsheets**, just attach/upload them to Hudu 


#### Question 6 - **Would you like to Convert PowerPoints to Articles?**
- **Convert powerpoints to articles**, effectively making an html table for this data
- **Don't convert powerpoints**, just attach/upload them to Hudu 


## Supported Files?

All files can be added if they are either
-under 100mb in size
-under 32000 Characters long when converted to html

Otherwise we just link to original file on SharePoint

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

All other non-document files skip conversion and are uploaded with a simple Hudu article containing file details.
<img width="723" height="45" alt="image" src="https://github.com/user-attachments/assets/619859ee-b367-483d-85ed-b2336f7eda34" />
most common formats that we skip conversion for are:
```
".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a", ".dll", ".so", ".lib", ".bin", ".class", ".pyc", ".pyo", ".o", ".obj", ".exe", ".msi", ".bat", ".cmd", ".sh", ".jar", ".app", ".apk", ".dmg", ".iso", ".img", ".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".tgz", ".lz", ".mp4", ".avi", ".mov", ".wmv", ".mkv", ".webm", ".flv", ".psd", ".ai", ".eps", ".indd", ".sketch", ".fig", ".xd", ".blend", ".ds_store", ".thumbs", ".lnk", ".heic"
```

Files that **don't have an extension**, we'll attempt to decode these as UTF-8

Files that aren't traditional documents will have a page generated for them with your chosen links. Here's what an image in sharepoint will look like after adding to Hudu: 
<img width="1184" height="1078" alt="image" src="https://github.com/user-attachments/assets/6a8fc563-49ee-4ed6-b17a-b35c78cc42f1" />

### If a file is too large after conversion (longer than 32000 characters or 100mb)
if this condition is met, the original is uploaded in Hudu and Linked
 <img width="849" height="363" alt="image" src="https://github.com/user-attachments/assets/4354122b-2f97-41ce-b64e-9e0c6262072e" />

#### Effectively, here's the process:
1. Select desired SharePoint Site(s)
2. Download all files from selected sites
3. Convert those that can be converted
4. Those that can be converted- Parse links, extract images, etc
5. Those that can't - generate a useful landing page for file


## FAQs:

#### Q: Does my original SharePoint Folder Structure Carry Over?
A: Yes. If you select a company to attribute an article to, the original folder structure is retained for that article
<img width="822" height="449" alt="image" src="https://github.com/user-attachments/assets/4e34671d-e769-4553-9566-e573ffb03720" />

#### Q: What does a PowerPoint file look like after converting into Hudu?
A: Pretty basic. Information from each slide is extracted, including images. This information is extracted into a web-friendly Article
<img width="896" height="498" alt="image" src="https://github.com/user-attachments/assets/3ba89439-61fb-4995-a46e-2fd2595b6ab7" />

#### Q: What does a Spreadsheet file look like after converting into Hudu?
A: It constructs an HTML table from your spreadsheet. Large spreadsheets are less manageable and any-sized spreadsheet will lose its formulae. It can be good for making sure information from SharePoint is searchable and findable in Hudu, however.
<img width="901" height="483" alt="image" src="https://github.com/user-attachments/assets/1d3188bc-495a-4553-8d81-ade5a349ee4c" />

## Community & Socials

[![Hudu Community](https://img.shields.io/badge/Community-Forum-blue?logo=discourse)](https://community.hudu.com/)
[![Reddit](https://img.shields.io/badge/Reddit-r%2Fhudu-FF4500?logo=reddit)](https://www.reddit.com/r/hudu)
[![YouTube](https://img.shields.io/badge/YouTube-Hudu-red?logo=youtube)](https://www.youtube.com/@hudu1715)
[![X (Twitter)](https://img.shields.io/badge/X-@HuduHQ-black?logo=x)](https://x.com/HuduHQ)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Hudu_Technologies-0A66C2?logo=linkedin)](https://www.linkedin.com/company/hudu-technologies/)
[![Facebook](https://img.shields.io/badge/Facebook-HuduHQ-1877F2?logo=facebook)](https://www.facebook.com/HuduHQ/)
[![Instagram](https://img.shields.io/badge/Instagram-@huduhq-E4405F?logo=instagram)](https://www.instagram.com/huduhq/)
[![Feature Requests](https://img.shields.io/badge/Feedback-Feature_Requests-brightgreen?logo=github)](https://hudu.canny.io/)

