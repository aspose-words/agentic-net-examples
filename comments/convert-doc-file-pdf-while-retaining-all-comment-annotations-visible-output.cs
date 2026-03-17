using System;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

class ConvertDocToPdfWithComments
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path to the destination PDF file.
        string outputPath = @"C:\Docs\ConvertedDocument.pdf";

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Make comments visible as PDF annotations.
        // ShowInAnnotations is only supported for PDF output.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

        // Rebuild the layout after changing the comment display mode.
        doc.UpdatePageLayout();

        // Create PDF save options (default constructor is sufficient).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF with the specified options.
        doc.Save(outputPath, pdfOptions);
    }
}
