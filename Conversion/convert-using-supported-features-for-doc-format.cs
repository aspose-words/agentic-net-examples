using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing document (any supported format).
        // The constructor automatically detects the format.
        Document doc = new Document("InputDocument.docx");

        // Create save options for the legacy DOC format.
        // The constructor overload that accepts a SaveFormat ensures the correct format is used.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Optional: set a password for the saved DOC file.
        // This does not encrypt the content, it only protects opening in Word.
        saveOptions.Password = "MyPassword";

        // Save the document as a binary DOC file using the specified options.
        doc.Save("OutputDocument.doc", saveOptions);
    }
}
