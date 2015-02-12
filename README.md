# Edit Document In Office
Open documents in MS Office for online editing via WebDav.

For Internet Explorer it uses ActiveXObject:
```
new ActiveXObject('SharePoint.OpenDocuments.3')
```

For other browsers - winFirefoxPlugin:
```
<object type="application/x-sharepoint" />
```

## Usage

```javascript
if (DocumentEditing.OfficeDocumentEditor.IsSupported())
{
    // Construct url to the document
    var url = "";

    DocumentEditing.OfficeDocumentEditor.EditDocument(url);
} 
```

## Requirements
* jQuery