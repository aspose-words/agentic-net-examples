using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content to generate multiple pages.
        for (int i = 0; i < 3; i++)
        {
            builder.Writeln($"This is page {i + 1}.");
            if (i < 2) // Insert a page break after each page except the last.
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Move the builder cursor to the primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        // Center the footer text.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Build the "Page X of Y" field sequence.
        builder.Write("Page ");
        builder.InsertField(FieldType.FieldPage, true);      // PAGE field.
        builder.Write(" of ");
        builder.InsertField(FieldType.FieldNumPages, true); // NUMPAGES field.

        // Update all fields so they display correct values.
        doc.UpdateFields();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PageNumberFooter.docx");
        doc.Save(outputPath);
    }
}
