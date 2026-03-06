using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDoc
{
    static void Main()
    {
        // Load an existing document (any supported format, e.g., DOCX).
        Document doc = new Document("Input.docx");

        // Option 1: Save directly using the SaveFormat enumeration.
        doc.Save("Output_Direct.doc", SaveFormat.Doc);

        // Option 2: Save using DocSaveOptions for additional control (e.g., password, routing slip).
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        // Example: set a password (optional).
        // saveOptions.Password = "MyPassword";
        doc.Save("Output_WithOptions.doc", saveOptions);
    }
}
