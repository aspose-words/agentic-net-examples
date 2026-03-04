using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ConvertToDocExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing document (any supported format, e.g., DOCX)
            string inputPath = @"C:\Docs\SourceDocument.docx";
            Document doc = new Document(inputPath);

            // Save the document in the legacy Microsoft Word 97‑2007 DOC format
            string outputPath = @"C:\Docs\ConvertedDocument.doc";
            doc.Save(outputPath, SaveFormat.Doc);

            // ------------------------------------------------------------
            // Alternative: use DocSaveOptions to control additional DOC options
            DocSaveOptions docOptions = new DocSaveOptions(SaveFormat.Doc)
            {
                // Example option: protect the saved file with a password
                Password = "SecretPassword",
                // Example option: preserve routing slip if present
                SaveRoutingSlip = true
            };

            string outputPathWithOptions = @"C:\Docs\ConvertedDocument_WithOptions.doc";
            doc.Save(outputPathWithOptions, docOptions);
        }
    }
}
