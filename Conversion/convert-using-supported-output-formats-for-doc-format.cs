using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDoc
{
    static void Main()
    {
        // Path to the source document (can be any supported format).
        string inputPath = @"C:\Input\SourceDocument.docx";

        // Path where the converted DOC file will be saved.
        string outputPathSimple = @"C:\Output\ConvertedSimple.doc";
        string outputPathWithOptions = @"C:\Output\ConvertedWithOptions.doc";

        // Load the document using the Document constructor (load rule).
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // Simple conversion: save directly to DOC format using SaveFormat.
        // -----------------------------------------------------------------
        doc.Save(outputPathSimple, SaveFormat.Doc);

        // -----------------------------------------------------------------
        // Conversion with save options: use DocSaveOptions for the DOC format.
        // -----------------------------------------------------------------
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        // Example of setting an optional password (uncomment if needed).
        // saveOptions.Password = "MyPassword";

        doc.Save(outputPathWithOptions, saveOptions);
    }
}
