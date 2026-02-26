using System;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add visible text.
        builder.Writeln("Visible text line 1.");

        // Add hidden text.
        builder.Font.Hidden = true;
        builder.Writeln("This text is hidden.");

        // Add another visible line.
        builder.Font.Hidden = false;
        builder.Writeln("Visible text line 2.");

        // Enable layout options.
        doc.LayoutOptions.ShowHiddenText = true;          // Render hidden text.
        doc.LayoutOptions.ShowParagraphMarks = true;     // Show paragraph marks (pilcrow).

        // Example: customize revision appearance.
        doc.LayoutOptions.RevisionOptions.InsertedTextColor = RevisionColor.BrightGreen;
        doc.LayoutOptions.RevisionOptions.ShowRevisionBars = false;

        // Rebuild the page layout after changing options.
        doc.UpdatePageLayout();

        // Save the document to PDF (or any other fixed-page format).
        doc.Save("RenderedDocument.pdf");
    }
}
