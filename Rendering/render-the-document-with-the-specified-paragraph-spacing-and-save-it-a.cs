using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set custom spacing before and after each paragraph (in points).
        builder.ParagraphFormat.SpaceBefore = 12; // 12 points before
        builder.ParagraphFormat.SpaceAfter = 12;  // 12 points after

        // Add sample paragraphs.
        builder.Writeln("First paragraph with custom spacing.");
        builder.Writeln("Second paragraph with custom spacing.");

        // Rebuild the page layout to ensure correct rendering.
        doc.UpdatePageLayout();

        // Save the document as a PDF file.
        doc.Save("RenderedDocument.pdf");
    }
}
