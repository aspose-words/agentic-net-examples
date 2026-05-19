using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace EvenPageFooterExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Set the built‑in Title property – this is the value the TITLE field will display.
            doc.BuiltInDocumentProperties.Title = "My Document Title";

            // Enable different footers for odd and even pages.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            // Move the builder cursor to the even‑page footer of the first (and only) section.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);

            // Align the paragraph to the left (default, but set explicitly for clarity).
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            // Insert a TITLE field that will display the document title.
            // Use the overload that inserts a field by type and updates it immediately.
            builder.InsertField(FieldType.FieldTitle, true);

            // Add some content to generate multiple pages so the even footer can be seen.
            builder.MoveToSection(0);
            builder.Writeln("Page 1 – odd page");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2 – even page (footer should appear here)");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3 – odd page");

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "EvenPageFooter.docx");
            doc.Save(outputPath);
        }
    }
}
