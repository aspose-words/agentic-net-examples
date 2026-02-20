using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class ConvertToDoc
{
    static void Main()
    {
        // Path to the source document (any supported format).
        string inputPath = @"C:\Docs\source.pdf";

        // Path where the DOC file will be saved.
        string outputPath = @"C:\Docs\converted.doc";

        // Load the document with default load options (auto-detect format).
        LoadOptions loadOptions = new LoadOptions();
        Document doc = new Document(inputPath, loadOptions);

        // Prepare save options for the DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions
        {
            // Explicitly set the target format (optional, but clarifies intent).
            SaveFormat = SaveFormat.Doc
        };

        // Save the document as a DOC file.
        doc.Save(outputPath, saveOptions);
    }
}
