// Load an existing document (any supported format) and save it as a legacy DOC file.
using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDoc
{
    static void Main()
    {
        // Path to the source document.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the DOC file will be saved.
        string outputPath = @"C:\Docs\ResultDocument.doc";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Create save options for the DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions
        {
            // Explicitly set the format to DOC (optional, as DocSaveOptions defaults to DOC).
            SaveFormat = SaveFormat.Doc
        };

        // Save the document using the specified options.
        doc.Save(outputPath, saveOptions);
    }
}
