using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class PdfContextualAccessExample
{
    static void Main()
    {
        // Load an existing PDF document with specific load options.
        var loadOptions = new PdfLoadOptions
        {
            // Example: do not skip images while loading.
            SkipPdfImages = false,
            // Load only the first page (optional).
            PageIndex = 0,
            PageCount = 1
        };

        // The Document constructor loads the file using the provided options.
        Document pdfDoc = new Document("Input.pdf", loadOptions);

        // Access a contextual member of the loaded document – for example, the total page count.
        int totalPages = pdfDoc.PageCount;
        Console.WriteLine($"Document contains {totalPages} page(s).");

        // Retrieve all shapes (including images) from the document.
        NodeCollection shapes = pdfDoc.GetChildNodes(NodeType.Shape, true);
        Console.WriteLine($"Document contains {shapes.Count} shape(s).");

        // Create PDF save options to customize the output PDF.
        var saveOptions = new PdfSaveOptions
        {
            // Show the document outline (bookmarks) when the PDF is opened.
            PageMode = PdfPageMode.UseOutlines,

            // Embed any OLE objects as annotations in the PDF.
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations,

            // Example: set compliance to PDF/A-1b.
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the document back to PDF using the configured options.
        pdfDoc.Save("Output.pdf", saveOptions);
    }
}
