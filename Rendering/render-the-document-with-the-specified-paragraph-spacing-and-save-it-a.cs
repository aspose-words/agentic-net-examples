using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set explicit spacing before and after each paragraph (in points).
        builder.ParagraphFormat.SpaceBefore = 12;   // 12 points before the paragraph.
        builder.ParagraphFormat.SpaceAfter = 12;    // 12 points after the paragraph.

        // Disable automatic spacing so the values above are used.
        builder.ParagraphFormat.SpaceBeforeAuto = false;
        builder.ParagraphFormat.SpaceAfterAuto = false;

        // Add a couple of paragraphs to demonstrate the spacing.
        builder.Writeln("First paragraph with custom spacing.");
        builder.Writeln("Second paragraph with the same custom spacing.");

        // Rebuild the page layout to ensure the spacing is taken into account when rendering.
        doc.UpdatePageLayout();

        // Save the document as a PDF file.
        doc.Save("Result.pdf", SaveFormat.Pdf);
    }
}
