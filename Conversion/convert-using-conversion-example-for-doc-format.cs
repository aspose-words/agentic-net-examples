using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDocExample
{
    static void Main()
    {
        // Path to the source document (can be any supported format, e.g., .docx, .pdf, .html, etc.).
        string inputPath = @"C:\Input\sample.docx";

        // Path where the converted DOC file will be saved.
        string outputPath = @"C:\Output\sample_converted.doc";

        // Load the source document into an Aspose.Words Document object.
        Document doc = new Document(inputPath);

        // Create save options for the DOC format.
        // DocSaveOptions allows additional settings specific to the legacy DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions
        {
            // Example: embed the generator name (default is true).
            ExportGeneratorName = true,

            // Example: compress all metafiles regardless of size.
            AlwaysCompressMetafiles = true
        };

        // Save the document in DOC format using the specified options.
        doc.Save(outputPath, saveOptions);

        Console.WriteLine("Document successfully converted to DOC format.");
    }
}
