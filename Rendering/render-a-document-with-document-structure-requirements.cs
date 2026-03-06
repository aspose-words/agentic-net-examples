using System;
using Aspose.Words;                       // Core document API
using Aspose.Words.Saving;                // PdfSaveOptions, ImageSaveOptions

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Build the document content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Add a heading using the built‑in "Heading 1" style.
        builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
        builder.Writeln("Document Title");
        // Add a normal paragraph.
        builder.ParagraphFormat.Style = doc.Styles["Normal"];
        builder.Writeln("This is a sample paragraph to demonstrate document structure export.");

        // Recalculate the page layout so that rendering works on the latest content.
        doc.UpdatePageLayout();

        // OPTIONAL: Render the first page to an image file for visual verification.
        // Using Aspose.Words' ImageSaveOptions avoids the need for System.Drawing.
        ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            PageSet = new PageSet(0) // Render only the first page.
        };
        doc.Save("RenderedPage.png", imgOptions);

        // Set PDF save options to export the document structure (tags).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true
        };

        // Save the document as PDF with the structure exported.
        doc.Save("DocumentWithStructure.pdf", pdfOptions);
    }
}
