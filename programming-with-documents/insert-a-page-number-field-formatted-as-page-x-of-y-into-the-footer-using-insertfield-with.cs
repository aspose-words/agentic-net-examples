using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsPageNumberFooter
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some content to generate multiple pages.
            builder.Writeln("First page content.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Second page content.");

            // Move the builder cursor to the primary footer.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            // Center the footer text.
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Insert the "Page X of Y" field sequence.
            builder.Write("Page ");
            // PAGE field – current page number.
            builder.InsertField(FieldType.FieldPage, true);
            builder.Write(" of ");
            // NUMPAGES field – total number of pages.
            builder.InsertField(FieldType.FieldNumPages, true);

            // Update all fields in the document to reflect the correct values.
            doc.UpdateFields();

            // Save the document to a file.
            doc.Save("PageNumberFooter.docx");
        }
    }
}
