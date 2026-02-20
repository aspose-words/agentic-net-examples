using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDoc
{
    static void Main()
    {
        // Input document path (any supported format, e.g., DOCX)
        string inputPath = @"C:\Input\sample.docx";

        // Output document path (DOC format)
        string outputPath = @"C:\Output\sample_converted.doc";

        // Load the source document using the default constructor (auto-detect format)
        Document doc = new Document(inputPath);

        // Create save options for the DOC format
        DocSaveOptions saveOptions = new DocSaveOptions
        {
            // Example: embed generator name (default true)
            ExportGeneratorName = true,
            // Example: compress all metafiles
            AlwaysCompressMetafiles = true
        };

        // Save the document as DOC using the specified options
        doc.Save(outputPath, saveOptions);

        Console.WriteLine("Conversion completed successfully.");
    }
}
