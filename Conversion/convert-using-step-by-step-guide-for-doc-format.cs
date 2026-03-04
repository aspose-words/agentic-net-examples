using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDoc
{
    static void Main()
    {
        // Step 1: Load the source document.
        // The constructor automatically detects the format of the input file.
        Document doc = new Document("InputDocument.docx");

        // Step 2: Save the document in the legacy DOC format using the Save method overload that accepts a SaveFormat.
        doc.Save("OutputDocument.doc", SaveFormat.Doc);

        // Alternative Step 2: Use DocSaveOptions for more control (e.g., setting a password).
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        // Example: protect the saved DOC with a password.
        saveOptions.Password = "MySecretPassword";

        // Save the document with the specified options.
        doc.Save("OutputDocumentWithPassword.doc", saveOptions);
    }
}
