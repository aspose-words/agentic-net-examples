using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source PDF document.
        Document pdfDocument = new Document("input.pdf");

        // Create save options for the DOC format.
        DocSaveOptions docSaveOptions = new DocSaveOptions
        {
            // Explicitly set the format to DOC (optional, DocSaveOptions defaults to DOC).
            SaveFormat = SaveFormat.Doc
        };

        // Save the loaded document as a DOC file using the specified options.
        pdfDocument.Save("output.doc", docSaveOptions);
    }
}
