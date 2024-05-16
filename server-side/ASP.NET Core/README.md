# EJ2 Document editor Services

## Available Web API services in EJ2 Document Editor
* [Import](./DefaultWebServicesReadMe.md#import)
* [SpellCheck](./DefaultWebServicesReadMe.md#spell-check)
* [SystemClipboard](./DefaultWebServicesReadMe.md#systemclipboard)
* [RestrictEditing](./DefaultWebServicesReadMe.md#restrictediting)
* [AutoSave(Version history helper)](#autosave)
* [LoadLatestVersionDocument(Version history helper)](#loadlatestversiondocument)
* [GetVersionData(Version history helper)](#getversiondata)
* [CompareSelectedVersion(Version history helper)](#compareselectedversion)
* [Download(Version history helper)](#download)


## AutoSave

Save the document along with the modified time and user from the client-side.
```
public void AutoSave([FromBody] ExportData exportData)
{
    byte[] data = Convert.FromBase64String(exportData.documentData.Split(',')[1]);
    string[] fileEntries = System.IO.Directory.GetFiles(Path.Combine("App_Data", exportData.fileName), "*.docx");
    string filePath = Path.Combine("App_Data", exportData.fileName, string.Format("v{0}.docx", (fileEntries.Length + 1)));
    using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
    {
        fs.Write(data, 0, data.Length);
    }
    // Read existing JSON data
    string jsonFilePath = Path.Combine("App_Data", "fileNameWithUserName.json");
    Dictionary<string, string> existingData;
    if (System.IO.File.Exists(jsonFilePath))
    {
        string existingJson = System.IO.File.ReadAllText(jsonFilePath);
        existingData = JsonConvert.DeserializeObject<Dictionary<string, string>>(existingJson);
    }
    else
    {
        existingData = new Dictionary<string, string>();
    }

    // Add the new data to the existing dictionary
    existingData[filePath] = exportData.modifiedUser;

    // Serialize the updated dictionary to JSON
    string updatedJson = JsonConvert.SerializeObject(existingData, Formatting.Indented);

    // Write the updated JSON data to the file
    System.IO.File.WriteAllText(jsonFilePath, updatedJson);
}
```

## LoadLatestVersionDocument

Retrieve the latest document version to open on client-side for editing.

```
public string LoadLatestVersionDocument([FromBody] UploadDocument doc)
{
    string[] fileEntries = System.IO.Directory.GetFiles("App_Data/" + doc.DocumentName, "*.docx");
    DirectoryInfo directoryInfo = new DirectoryInfo("App_Data/" + doc.DocumentName);

    // Get all files in the directory
    FileInfo[] files = directoryInfo.GetFiles();

    // Get the last modified file
    FileInfo lastModifiedFile = files
        .OrderByDescending(f => f.LastWriteTime)
        .FirstOrDefault();


    Stream stream = System.IO.File.OpenRead(Path.Combine("App_Data", doc.DocumentName, lastModifiedFile.Name));
    stream.Position = 0;

    WordDocument document = WordDocument.Load(stream, FormatType.Docx);
    string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
    document.Dispose();
    return json;
}
```

## GetVersionData

Retrieve the saved document version from server-side for listing on the client-side.
```
public string GetVersionData([FromBody] UploadDocument doc)
{
    string directoryPath = "App_Data/" + doc.DocumentName;
    string[] fileEntries = System.IO.Directory.GetFiles("App_Data/" + doc.DocumentName, "*.docx");
    List<Version> versions = new List<Version>();
    // Check if the directory exists
    if (Directory.Exists(directoryPath))
    {
        // Get files in the directory
        string[] files = Directory.GetFiles(directoryPath);

        // Display file names with their last write time
        foreach (string file in files)
        {
            string modifiedUser = "";
            if (System.IO.File.Exists("App_Data\\fileNameWithUserName.json"))
            {
                // Read existing JSON data of file name with user
                string existingJson = System.IO.File.ReadAllText("App_Data/fileNameWithUserName.json");

                // Deserialize JSON data into a Dictionary<string, string>
                Dictionary<string, string> data = JsonConvert.DeserializeObject<Dictionary<string, string>>(existingJson);

                string fileToCheck = file.Replace("/", "\\");
                // Check if the key of file exists in the dictionary
                if (data.ContainsKey(fileToCheck))
                {
                    // Key exists, retrieve its vale of user name
                    modifiedUser = data[fileToCheck];
                }
            }
            // Get the last write time of the file
            DateTime lastWriteTime = System.IO.File.GetLastWriteTime(file);
            versions.Add(new Version(System.IO.Path.GetFileName(file), file, modifiedUser, lastWriteTime));
        }
        var versionDates = versions.OrderByDescending(v => v.LastSavedTime).ToList();
        CompareData compare = new CompareData();
        if (versions.Count > 1)
        {
            // Compare versionDates[0] and versionDates[1]
            compare.Document = Compare(versionDates[0], versionDates[1]);
        }
        else
        {
            // Handle the case where version[1] doesn't exist
            FileStream originalDocumentStreamPath = new FileStream(versions[0].FullName, FileMode.Open, FileAccess.Read);
            compare.Document = WordDocument.Load(originalDocumentStreamPath, FormatType.Docx);
        }
        var groupedDates = versions.GroupBy(version => version.LastSavedTime.Date).OrderByDescending(group => group.Key); ;
        // Iterate over grouped dates and create RootObject instances
        List<RootObject> rootObjects = new List<RootObject>();
        foreach (var group in groupedDates)
        {

            List<SubChild> childObjects = new List<SubChild>();
            List<Version> tempVersion = group.ToArray().OrderByDescending(v => v.LastSavedTime).ToList();
            foreach (var chid in tempVersion)
            {
                SubChild childObject = new SubChild
                {
                    id = chid.DocumentVersion,
                    name = chid.LastSavedTime.ToString("MMMM dd, hh:mm tt"),
                    user = chid.ModifiedUser

                };
                childObjects.Add(childObject);
            }
            // Create RootObject instance for each group
            RootObject rootObject = new RootObject
            {
                id = group.Key.ToString("yyyy-MM-dd"),
                name = group.Key.ToString("MMMM dd"),
                user = childObjects[0].user,
                subChild = childObjects
            };

            rootObjects.Add(rootObject);
        }
        compare.Data = rootObjects;
        return Newtonsoft.Json.JsonConvert.SerializeObject(compare);
    }
    return null;
}
```

## CompareSelectedVersion
Compare the selected revision with previous revision and show the changes in tracked content.
```
public string CompareSelectedVersion([FromBody] CompareDocument doc)
{
    string[] fileEntries = System.IO.Directory.GetFiles("App_Data/" + doc.DocumentName, "*.docx");

    string targetFileName = doc.SelectedVersion;
    string previousFile = null;
    bool foundTargetFile = false;
    List<Version> versions = new List<Version>();
    string sfdtString = "";

    foreach (string filePath in fileEntries)
    {
        if (!foundTargetFile)
        {
            if (System.IO.Path.GetFileName(filePath) == targetFileName)
            {
                versions.Add(new Version(System.IO.Path.GetFileName(filePath), filePath, null));
                foundTargetFile = true;
            }
            else
            {
                previousFile = filePath; // Store the previous file
            }
        }
        else
        {
            break; // Exit the loop after finding the target file
        }
    }

    if (foundTargetFile && previousFile != null)
    {
        versions.Add(new Version(System.IO.Path.GetFileName(previousFile), previousFile, null, null));
        WordDocument result = Compare(versions[0], versions[1]);
        sfdtString = JsonConvert.SerializeObject(result);
    }
    else
    {
        FileStream originalDocumentStreamPath = new FileStream(versions[0].FullName, FileMode.Open, FileAccess.Read);
        WordDocument document = WordDocument.Load(originalDocumentStreamPath, FormatType.Docx);
        sfdtString = JsonConvert.SerializeObject(document);
        document.Dispose();
    }
    return sfdtString;
}
//Compare the word document using DocIO https://help.syncfusion.com/file-formats/docio/word-document/compare-word-documents
private WordDocument Compare(Version o, Version n)
{
    using (FileStream originalDocumentStreamPath = new FileStream(o.FullName, FileMode.Open, FileAccess.Read))
    {
        using (WDocument originalDocument = new WDocument(originalDocumentStreamPath, WFormatType.Docx))
        {
            //Load the revised document.
            using (FileStream revisedDocumentStreamPath = new FileStream(n.FullName, FileMode.Open, FileAccess.Read))
            {
                using (WDocument revisedDocument = new WDocument(revisedDocumentStreamPath, WFormatType.Docx))
                {
                    // Compare the original and revised Word documents.
                    originalDocument.Compare(revisedDocument);

                    return WordDocument.Load(originalDocument);
                }
            }
        }
    }
}
```

## Download
Download the selected version as a copy.
```
public IActionResult Download([FromBody] DownloadData data)
{
    // Specify the file path relative to the App_Data folder
    string directoryPath = Path.Combine(_hostingEnvironment.ContentRootPath, "App_Data");
    string folderName = data.DocumentName;
    string versionFolder = data.SelectedVersion;
    string filePath = Path.Combine(directoryPath, folderName, versionFolder);

    // Check if the file exists
    if (!System.IO.File.Exists(filePath))
    {
        return NotFound(); // Return 404 Not Found if file doesn't exist
    }

    // Open the file stream
    FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
    string fileName = data.DocumentName.Split('.')[0] + "_" + data.SelectedVersion;
    // Return FileStreamResult with the file stream
    return new FileStreamResult(fileStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    {
        FileDownloadName = fileName // Specify the file name to be used when downloaded
    };
}
```

