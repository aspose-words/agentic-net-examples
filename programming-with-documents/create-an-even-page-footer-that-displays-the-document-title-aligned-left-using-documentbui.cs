using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the document title – this will be used by the TITLE field.
        doc.BuiltInDocumentProperties.Title = "Sample Document Title";

        // Configure the page setup to have different footers for odd and even pages.
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Add enough text to generate at least two pages.
        builder.Font.Size = 12;
        for (int i = 0; i < 100; i++)
        {
            builder.Writeln($"Line {i + 1}: This is sample text to fill the document and create multiple pages.");
        }

        // Move to the even-page footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);

        // Align the paragraph to the left.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

        // Insert the TITLE field which displays the document title.
        builder.InsertField("TITLE", "");

        // Save the document.
        string outputPath = "EvenPageFooter.docx";
        doc.Save(outputPath);
    }
}
