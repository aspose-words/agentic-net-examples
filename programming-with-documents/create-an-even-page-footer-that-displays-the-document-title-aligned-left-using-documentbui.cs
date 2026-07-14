using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

namespace AsposeWordsEvenFooterExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Set the built‑in Title property – this is the value the TITLE field will display.
            doc.BuiltInDocumentProperties.Title = "Sample Document Title";

            // Create a DocumentBuilder to edit the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Enable different footers for odd and even pages.
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            // Move the builder cursor to the even‑page footer.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);

            // Align the footer text to the left.
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            // Insert a TITLE field that will display the document title.
            // The field is updated immediately (second argument = true).
            builder.InsertField(FieldType.FieldTitle, true);

            // Add some body content to generate multiple pages.
            builder.MoveToSection(0);
            builder.Writeln("Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3");

            // Save the document to disk.
            doc.Save("EvenFooterTitle.docx");
        }
    }
}
