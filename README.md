# 📂 VBA Excel Form Template for Data Entry & SharePoint Integration  

## 📌 Overview  

This **Excel VBA-based form** was created for a simulated use case where users can **enter structured data, select a PDF file, and store everything in an Excel sheet while managing files in a specified folder (local or SharePoint).**  

Here is how the main Overview looks, featuring two buttons at the top and displaying the previously entered information, including links to the associated files:
![image](https://github.com/user-attachments/assets/b1fda2bc-065e-4100-b422-4c10efda7e97)
And this is the form that opens after clicking the 'New Entry' button:

<img src="https://github.com/user-attachments/assets/639f123e-241e-4ff0-bed7-cdaa55b01218" height="400">

---

## 🛠 Features & Workflow  

### 1️⃣ Data Entry Process  

1. Click **"NEW entry"** → Opens a user-friendly VBA form.  
2. Enter key details:  
   - **Title** → Free text field (must be unique).  
   - **Date** → Selected from a date field.  
   - **Duration** → Number input (years).  
   - **Category** → Selected from a dropdown.  
   - **Organisation** → Selected from a dropdown (option to add new).  
   - **Project Lead** → Multi-selection list (option to add new).  
3. Click **"Choose your PDF file & SAVE"** →  
   - The selected **PDF file is moved** from its original location to a **designated folder** (local or SharePoint).  
   - **All entered data is recorded** in the Excel sheet.  
4. Click **"Open SharePoint"** to directly access the folder with stored PDFs.  
5. Each saved entry in Excel has a **hyperlinked file path**, allowing users to open the specific document with one click.  

---

## 2️⃣ File Handling Logic  

- **The selected PDF is automatically moved** to a specified storage location.  
- **Original file is deleted from its initial folder** to ensure a single source of truth.  
- The target folder can be either:  
  - ✅ **A local directory** (as shown in the example).  
  - ✅ **A SharePoint folder** (requires integration—see details below).  

---

## 3️⃣ Duplicate Entry Prevention  

- The VBA script **checks for duplicate Titles** before saving data.  
- If a duplicate is found, an error message appears, preventing redundant entries.  
- This feature is useful for real-world applications where unique record-keeping is required (e.g., project tracking, contract management).  

---

## 4️⃣ Customization Options  

- **Form Labels** → Feel free to rename labels in the form as needed.  
- **Input Fields** → The form includes a variety of field types to demonstrate versatility:  
  - **Text field** → Title  
  - **Date picker** → Date  
  - **Number input** → Duration  
  - **Dropdown lists** → Category, Organisation  
  - **Multi-selection list** → Project Lead (supports adding new entries)  
- **Sheet Protection** →  
  - To **prevent accidental edits**, consider **protecting the sheet** while allowing VBA modifications.  
  - **Important:** If the sheet is protected, don’t forget to **add a password in the VBA code** for unlocking before data entry.  

---

## 📁 How to Save Files in SharePoint Instead of Local Folder  

If you want to store files in **SharePoint instead of a local folder**, follow these modifications:  

### 🔹 Find the SharePoint Document Library URL  

1. Open the SharePoint folder where you want to save files.  
2. Copy the full **URL path** to the document library.  

### 🔹 Update VBA Code to Use SharePoint Path  

Modify the file move logic in VBA:  

```vba
'UserForm:
If Not (CheckSharedDrive(driveLetter)) Then
  MapSharepointToDrive (driveLetter)
End If

reName = "https://yourcompany.sharepoint.com/sites/YourSite/Shared Documents/YourFolder/" & title & ".pdf"
NameExcel = "https://yourcompany.sharepoint.com/sites/YourSite/Shared Documents/YourFolder/" & title & ".pdf"

'Store the User Name
'  AddedBy = Application.UserName
'  With Sheets("Overview").Cells(65536, 7).End(xlUp)
'    .Offset(1, 0) = AddedBy
'    .Font.Name = "RWE Sans"
'  End With

sharepointFolder = "https://yourcompany.sharepoint.com/sites/YourSite/Shared Documents/YourFolder/"

'Modul2:
x = Shell("C:\Windows\Explorer.exe /n,/e,""A:/YourSite/Shared Documents/YourFolder/""", vbNormalFocus)
