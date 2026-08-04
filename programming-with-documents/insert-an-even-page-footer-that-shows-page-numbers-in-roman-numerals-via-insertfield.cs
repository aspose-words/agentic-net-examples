using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different footers for odd and even pages.
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Add a few pages so that we have both odd and even pages.
        for (int i = 1; i <= 4; i++)
        {
            builder.Writeln($"Content of page {i}");
            if (i < 4)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Set the page number style for the whole section to uppercase Roman numerals.
        doc.FirstSection.PageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;

        // Move the builder cursor to the even-page footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);

        // Insert a PAGE field that will display the page number.
        builder.Write("Page ");
        builder.InsertField("PAGE", "");

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "EvenFooterRoman.docx");
        doc.Save(outputPath);
    }
}
