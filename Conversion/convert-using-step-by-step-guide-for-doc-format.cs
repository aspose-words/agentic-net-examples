using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDoc
{
    static void Main()
    {
        // Path to the source document (any supported format).
        string sourcePath = @"C:\Docs\source.docx";

        // Load the document. The format is detected automatically.
        Document doc = new Document(sourcePath);

        // Configure save options for the legacy DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions
        {
            // Explicitly set the target format (optional, DocSaveOptions defaults to DOC).
            SaveFormat = SaveFormat.Doc,

            // Example option: embed the Aspose.Words generator name.
            ExportGeneratorName = true,

            // Example option: compress all metafiles.
            AlwaysCompressMetafiles = true
        };

        // Path for the output DOC file.
        string outputPath = @"C:\Docs\output.doc";

        // Save the document in DOC format using the specified options.
        doc.Save(outputPath, saveOptions);
    }
}
