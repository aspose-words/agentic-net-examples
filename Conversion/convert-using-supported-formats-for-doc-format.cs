using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDoc
{
    static void Main()
    {
        // Path to the source document (can be any supported format, e.g., .docx, .pdf, .html, etc.)
        string sourcePath = @"C:\Input\SampleDocument.docx";

        // Path where the converted .doc file will be saved.
        string outputPath = @"C:\Output\SampleDocument.doc";

        // Load the source document. The constructor automatically detects the format.
        Document doc = new Document(sourcePath);

        // Create DocSaveOptions specifying the DOC format.
        // This uses the provided DocSaveOptions constructor that accepts a SaveFormat.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Optional: set additional options, e.g., a password or routing slip preservation.
        // saveOptions.Password = "MyPassword";
        // saveOptions.SaveRoutingSlip = true;

        // Save the document in the legacy DOC format using the save options.
        doc.Save(outputPath, saveOptions);

        // Inform the user that conversion is complete.
        Console.WriteLine($"Document converted and saved to: {outputPath}");
    }
}
