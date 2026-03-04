using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDoc
{
    static void Main()
    {
        // Load an existing document (any supported format, e.g., DOCX).
        Document doc = new Document("InputDocument.docx");

        // Create save options for the legacy DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Optional: set a password for the saved DOC file.
        // saveOptions.Password = "SecretPassword";

        // Save the document as a .doc file using the specified options.
        doc.Save("OutputDocument.doc", saveOptions);
    }
}
