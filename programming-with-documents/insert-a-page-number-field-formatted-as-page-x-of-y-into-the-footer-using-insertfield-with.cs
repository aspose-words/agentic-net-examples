using System;
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
            if (i < 2)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Move the builder to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        // Center the footer text.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Insert the page number field in the format "Page X of Y".
        builder.Write("Page ");
        // PAGE field with Arabic numeral switch.
        builder.InsertField("PAGE \\* Arabic", "");
        builder.Write(" of ");
        // NUMPAGES field with Arabic numeral switch.
        builder.InsertField("NUMPAGES \\* Arabic", "");

        // Update all fields so they display the correct values.
        doc.UpdateFields();

        // Save the document to the local file system.
        doc.Save("PageNumberFooter.docx");
    }
}
