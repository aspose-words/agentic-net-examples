using System;
using System.IO;
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

            // Set the document title – this will be displayed by the TITLE field.
            doc.BuiltInDocumentProperties.Title = "Sample Document Title";

            // Enable different footers for odd and even pages.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            // Move the builder cursor to the even-page footer.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);

            // Ensure left alignment (default, but set explicitly for clarity).
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            // Insert a TITLE field that will display the document title.
            // InsertField(FieldType, bool) with updateField = false, then update manually.
            FieldTitle titleField = (FieldTitle)builder.InsertField(FieldType.FieldTitle, false);
            titleField.Update();

            // Add some pages to the document so that the even footer can be seen.
            builder.MoveToSection(0);
            builder.Writeln("Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3");

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "EvenFooter.docx");
            doc.Save(outputPath);
        }
    }
}
