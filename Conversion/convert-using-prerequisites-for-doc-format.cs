using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source document (any format supported by Aspose.Words).
        string inputPath = @"C:\Docs\SampleDocument.docx";

        // Path for the converted DOC file.
        string outputPath = @"C:\Docs\ConvertedDocument.doc";

        // Load the source document.
        Document doc = new Document(inputPath);

        // Create save options for the legacy DOC format.
        // The constructor that accepts a SaveFormat ensures the correct format is set.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Optional: set a password for the saved DOC file.
        // saveOptions.Password = "MyPassword";

        // Save the document as DOC using the specified options.
        doc.Save(outputPath, saveOptions);
    }
}
