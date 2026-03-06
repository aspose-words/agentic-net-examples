using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing document (any supported format, e.g., DOCX).
        Document doc = new Document("InputDocument.docx");

        // Create save options for the legacy DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Optional: set a password to protect the saved DOC file.
        saveOptions.Password = "MySecretPassword";

        // Save the document as a DOC file using the specified options.
        doc.Save("ConvertedDocument.doc", saveOptions);
    }
}
