# Hudu SharePoint Migration
Easy Migration from SharePoint to Hudu

## Note on Intended Purpose
If you are working from a single mapped SharePoint drive rather than multiple SharePoint sites, you may prefer the direct [Files-Hudu-Migration](https://github.com/Hudu-Technologies-Inc/Files-Hudu-Migration) tool. However, for environments with multiple SharePoint sites, this project is the recommended solution. It preserves inter-site links and treats each SharePoint site as a distinct entity, whereas Files-Hudu-Migration processes each file or folder as an individual entity.

### Prerequisites

- Hudu Instance of 2.42.0 or newer
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

### Optional Environment Settings

The migration can be tuned by uncommenting or adding settings in your environment file before the final `. .\Sharepoint-Migration.ps1` line.

#### Resume and Disk Usage

Use resume mode to skip SharePoint drive items that were already completed in a prior run. By default, the script compares SharePoint item ID and ETag from `logs\sharepoint-migration-state.jsonl`.

```powershell
$SharePointResumeFromState = $true
$SharePointResumeIgnoreETag = $false
```

Low-disk mode processes one site/drive/root item batch at a time and clears working files after each batch.

```powershell
$SharePointLowDiskMode = $true
```

To avoid creating duplicate Hudu articles, the script can skip articles when the exact title already exists in the target company or global KB.

```powershell
$SharePointSkipExistingArticles = $true
```

When the destination company can be resolved without prompting, this check runs immediately after SharePoint discovery so matching files are removed before conversion, indexing, stubbing, image upload, or article population. Skipped files are still written to the resume state, so reinvoking the migration does not reprocess them.

#### Site Selection Filters

You can hide known-unwanted SharePoint sites from the selection prompts and from "all sites" runs. Matching is case/punctuation-insensitive and checks both SharePoint `displayName` and `name`.

```powershell
$SharePointSiteSkipNames = @(
  "Archive",
  "Old Client Portal"
)
```

#### File Conversion Controls

Some file types can be grouped into folder-level Hudu file index articles instead of becoming one article per file.

```powershell
$SharePointIndexOnlyExtensions = @(".eps", ".ai", ".psd", ".indd")
```

PDFs can also be attached to generated stub articles instead of converted to HTML.

```powershell
$SharePointPdfUploadAsFile = $true
```

Paths containing literal wildcard characters, such as `[NEW CLIENT]`, are handled with literal path checks during conversion, reading, and attachment upload.

#### Client Attribution

Client attribution can build a SharePoint-client-to-Hudu-company map from either a local `clients.json` file or from one or more SharePoint lists. Matching normalizes case, accents, punctuation, and legal suffixes where practical, then uses high-confidence exact/compact matches before fuzzy matching.

```powershell
$SharePointClientAttributionEnabled = $true
$SharePointClientAttributionAutoApply = $true
$SharePointClientAttributionClientsPath = ".\clients.json"
$SharePointClientAttributionListNames = @("Client List")
```

`clients.json` may be a simple string array:

```json
[
  "Example Company (EXAMPLE) [Provider]"
  "Exemplary Co (EXAMPLE) [Provider]"
]
```

Or it may contain richer objects with aliases or known Hudu company IDs:

```json
[
  {
    "name": "Jolly Inc.",
    "aliases": ["Jolly"],
    "huduCompanyId": 123
  }
]
```

If your structured SharePoint lists use a client picker or lookup column, configure the field names here. SharePoint internal names such as `Select_x0020_a_x0020_Client` are decoded before matching.

```powershell
$SharePointClientAttributionFieldNames = @(
  "Select a Client",
  "Client Name",
  "Client",
  "Customer",
  "Company",
  "LinkTitle"
)
```

When using the destination option "Match/Create one company per SharePoint site", client attribution is tried before site-name attribution by default. This helps prevent generic sites like `Calendar`, `Export`, or `Management` from becoming the company assignment when document/list metadata contains a better client signal.

```powershell
$SharePointPreferClientAttributionOverSiteCompany = $true
```

The script can also roll up list item metadata into per-site and per-list client designations. For example, if most items in a list have `Select a Client` set to `Jolly`, the list can inherit Jolly as its default company. If most matching client fields across a site point to the same company, documents in that site can inherit the site designation before falling back to broader path/title matching.

```powershell
$SharePointClientAttributionUseSiteDesignations = $true
$SharePointClientAttributionUseListDesignations = $true
$SharePointClientAttributionDesignationMinShare = 0.8
$SharePointClientAttributionDesignationMinItems = 1
```

Designation maps are written to `logs\client-designation-map.json` for review. Raise `SharePointClientAttributionDesignationMinShare` for mixed-client sites/lists; lower it only when the source metadata is sparse but trustworthy.

These thresholds control how confident matches need to be before auto-application. Raising them reduces false positives; lowering them increases automation.

```powershell
$SharePointClientAttributionMinScore = 95
$SharePointClientAttributionMinGap = 5
$SharePointClientAttributionListItemMinScore = 95
$SharePointClientAttributionListItemMinGap = 3
$SharePointClientAttributionCreateMissing = $false
```

Client attribution maps are cached in `logs\client-attribution-map.json`. If `clients.json` is newer than the cache, the map is rebuilt automatically. You can force a rebuild:

```powershell
$SharePointClientAttributionUseCachedMap = $true
$SharePointClientAttributionForceRebuildMap = $true
```

#### Site Company Attribution

The per-site destination mode can match SharePoint site names to Hudu companies, optionally creating missing companies. This is useful when each SharePoint site really represents one client. If client attribution is preferred, unmatched site names will not create companies automatically.

```powershell
$SharePointSiteCompanyMinScore = 95
$SharePointSiteCompanyMinGap = 5
$SharePointSiteCompanyCreateMissing = $true
$SharePointSiteCompanyUseCachedMap = $true
$SharePointSiteCompanyForceRebuildMap = $false
```

#### Structured List JSON Export

You can export selected SharePoint lists as per-company JSON bundles for later asset import. This is useful for inventory-like lists such as network devices, printers, contacts, locations, and ISP info.

```powershell
$SharePointStructuredListJsonNames = @(
  "Mapped Drives",
  "Printers",
  "Network Devices",
  "Locations",
  "ISP Info"
)
```

To skip structured-list export and migrate files only, leave it unset or set it to an empty array:

```powershell
$SharePointStructuredListJsonNames = @()
```

To generate only the structured-list JSON bundles and stop before file conversion/article upload:

```powershell
$SharePointStructuredListJsonOnly = $true
```

#### Site Page Fetching

For an initial SharePoint site page export, enable the fetch job:

```powershell
$SharePointFetchSitePages = $true
```

The main runner will fetch pages after source site selection. You can also run the job manually from an authenticated session with `$userSelectedSites` populated:

```powershell
. .\jobs\Get-SitePages.ps1
```

This writes page HTML snapshots to `logs\site-pages-html`, raw page/webpart JSON to `logs\site-pages-json`, and a review CSV to `logs\site-pages-index.csv`. It uses Microsoft Graph site pages and webparts endpoints and is intended as a first-pass capture before wiring site pages into Hudu article creation.

To import fetched site page snapshots as Hudu articles, enable:

```powershell
$SharePointImportSitePagesAsArticles = $true
```

The importer treats site pages as pre-converted HTML documents and reuses the existing attribution, skip-existing, stub, populate, upload, and relink stages. Base64-embedded images are extracted to `logs\site-pages-assets` and queued as Hudu uploads; external image URLs are left in the HTML and recorded for review.

If you need to rerun page imports after improving page rendering, ignore only the site-page resume state with:

```powershell
$SharePointForceReimportSitePages = $true
```

This is separate from `$SharePointSkipExistingArticles`, which controls the duplicate Hudu article title/org check.

#### External Article Images

After importing pages, you can scan Hudu articles for absolute external `<img src="...">` values, upload those images to Hudu, and rewrite the image sources. Before downloading a remote image, the job can check existing Hudu uploads and public photos by exact filename and reuse a matching Hudu URL. The job is dry-run by default:

```powershell
$HuduInternalizeExternalArticleImagesDryRun = $true
. .\jobs\Internalize-ExternalArticleImages.ps1
```

Review `logs\internalized-external-images\internalized-external-images.csv`, then run with dry-run disabled when ready:

```powershell
$HuduInternalizeExternalArticleImagesDryRun = $false
. .\jobs\Internalize-ExternalArticleImages.ps1
```

The report also classifies unexpected local image sources. Expected Hudu image/file paths include relative or absolute `public_photo`, `public_photos`, `photo`, `photos`, `upload`, `uploads`, `file`, and `files` URLs, with or without a leading slash. To remove unexpected local/absolute image tags while internalizing external images, enable:

```powershell
$HuduInternalizeExternalArticleImagesScrubUnexpectedLocalSources = $true
```

To disable reuse of existing Hudu uploads/public photos before downloading external images, set:

```powershell
$HuduInternalizeExternalArticleImagesPreferExistingHuduImages = $false
```

Unexpected local/relative image sources are reported but not rewritten by default, even if a filename matches an existing Hudu upload/public photo. To allow those existing-Hudu rewrites in a separate cleanup pass, set:

```powershell
$HuduInternalizeExternalArticleImagesRewriteUnexpectedLocalExisting = $true
```

To check whether external images appear downloadable while staying in dry-run, enable the probe option. This sends a lightweight `HEAD` request first and falls back to a one-byte ranged `GET` when needed:

```powershell
$HuduInternalizeExternalArticleImagesDryRun = $true
$HuduInternalizeExternalArticleImagesProbeDownloads = $true
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
- transfer everything **to companies in Hudu by matching/creating one company per SharePoint site**

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
-under 196000 Characters long when converted to html

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

### If a file is too large after conversion (longer than 196000 characters or 100mb)
if this condition is met, the original is uploaded in Hudu and Linked
 <img width="849" height="363" alt="image" src="https://github.com/user-attachments/assets/4354122b-2f97-41ce-b64e-9e0c6262072e" />

#### Effectively, here's the process:
1. Select desired SharePoint Site(s)
2. Download all files from selected sites
3. Build optional attribution maps for SharePoint clients and/or sites
4. Optionally export selected structured SharePoint lists as JSON bundles
5. Convert files that can be converted
6. Parse converted HTML for links and embedded files
7. Generate useful landing pages for files that are not converted
8. Create, populate, and relink Hudu articles


## FAQs:

#### Q: Does my original SharePoint Folder Structure Carry Over?
A: Yes. If you select a company to attribute an article to, the original folder structure is retained for that article
<img width="822" height="449" alt="image" src="https://github.com/user-attachments/assets/4e34671d-e769-4553-9566-e573ffb03720" />

#### Q: Can I skip structured SharePoint data and migrate files only?
A: Yes. Set `$SharePointStructuredListJsonNames = @()` in your environment file.

#### Q: Can SharePoint list metadata decide which Hudu company an item belongs to?
A: Yes. Use `$SharePointClientAttributionListNames` to identify the SharePoint list that contains client names, and `$SharePointClientAttributionFieldNames` to identify fields on other lists that point to those clients. The matcher ignores common punctuation/case/accent differences and caches repeated lookups.

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
