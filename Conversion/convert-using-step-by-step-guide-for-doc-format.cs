using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDoc
{
    static void Main()
    {
        // Define input and output file paths.
        string inputPath = @"C:\Docs\SampleDocument.docx";
        string outputPath = @"C:\Docs\SampleDocumentConverted.doc";

        // Load the source document (any supported format, e.g., DOCX).
        Document doc = new Document(inputPath);

        // Option 1: Save directly using the SaveFormat enumeration.
        doc.Save(outputPath, SaveFormat.Doc);

        // Option 2: Use DocSaveOptions for more control (e.g., password protection).
        // Uncomment the following lines to use DocSaveOptions instead of the direct overload.
        /*
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        // Example: set a password (optional).
        // saveOptions.Password = "MySecretPassword";
        doc.Save(outputPath, saveOptions);
        */
    }
}
