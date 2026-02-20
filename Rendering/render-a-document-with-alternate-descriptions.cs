using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Layout;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Create PDF save options.
        PdfSaveOptions options = new PdfSaveOptions();

        // Render DrawingML shapes using their fallback representations.
        options.DmlRenderingMode = DmlRenderingMode.Fallback;

        // Show comments as PDF annotations instead of balloons.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

        // Rebuild the layout after changing the comment display mode.
        doc.UpdatePageLayout();

        // Save the document to PDF with the configured options.
        doc.Save("Output_AlternateDescriptions.pdf", options);
    }
}
