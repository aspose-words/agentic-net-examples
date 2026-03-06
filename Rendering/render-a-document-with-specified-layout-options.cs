using System;
using Aspose.Words;
using Aspose.Words.Layout;

class RenderDocumentWithLayoutOptions
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add normal text.
        builder.Writeln("This text is visible.");

        // Add hidden text.
        builder.Font.Hidden = true;
        builder.Writeln("This text is hidden.");

        // Add another paragraph to demonstrate paragraph marks.
        builder.Writeln("Another visible paragraph.");

        // Configure layout options.
        // Show hidden text in the rendered output.
        doc.LayoutOptions.ShowHiddenText = true;

        // Show paragraph marks (pilcrow) at the end of each paragraph.
        doc.LayoutOptions.ShowParagraphMarks = true;

        // Example of ignoring printer metrics (optional).
        // doc.LayoutOptions.IgnorePrinterMetrics = false;

        // Rebuild the page layout so that the changes take effect.
        doc.UpdatePageLayout();

        // Save the document to PDF. The file name determines the format.
        doc.Save("RenderedDocument.pdf");
    }
}
