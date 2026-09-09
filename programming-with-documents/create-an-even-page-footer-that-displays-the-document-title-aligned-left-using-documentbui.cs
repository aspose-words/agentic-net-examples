using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;

namespace AsposeWordsEvenFooterExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Set the document title – this value will be displayed by the TITLE field.
            doc.BuiltInDocumentProperties.Title = "Sample Document Title";

            // Use DocumentBuilder to add content and configure page setup.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Enable different footers for odd and even pages.
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            // Create an even‑page footer.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);

            // Align the footer text to the left.
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            // Insert a TITLE field that displays the document title.
            // The field is updated immediately.
            builder.InsertField(FieldType.FieldTitle, true);

            // Return to the main body of the document.
            builder.MoveToSection(0);
            builder.Writeln("Page 1 – odd page.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2 – even page (footer shows title).");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3 – odd page.");

            // Save the document to the local file system.
            doc.Save("EvenPageFooter.docx");
        }
    }
}
