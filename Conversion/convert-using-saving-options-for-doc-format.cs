using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDocWithOptions
{
    static void Main()
    {
        // Path to the source document (any format supported by Aspose.Words).
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the DOC file will be saved.
        string outputPath = @"C:\Docs\ConvertedDocument.doc";

        // Load the source document.
        Document doc = new Document(inputPath);

        // Create save options for the older DOC format.
        // The constructor that accepts a SaveFormat ensures the correct format is set.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Optional: set a password that will protect the saved DOC file.
        saveOptions.Password = "MyPassword";

        // Optional: preserve routing slip information if present.
        saveOptions.SaveRoutingSlip = true;

        // Save the document using the specified options.
        doc.Save(outputPath, saveOptions);
    }
}
