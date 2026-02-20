using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDoc
{
    static void Main()
    {
        // Load an existing document (any supported format) from disk.
        // The LoadFormat is detected automatically.
        Document doc = new Document("InputDocument.docx");

        // Create save options for the DOC format.
        // The SaveFormat property ensures the document is saved as a legacy .doc file.
        DocSaveOptions saveOptions = new DocSaveOptions
        {
            SaveFormat = SaveFormat.Doc
        };

        // Save the document in the Microsoft Word 97‑2007 .doc format.
        doc.Save("OutputDocument.doc", saveOptions);
    }
}
