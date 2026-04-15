using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

namespace AsposeWordsEvenFooter
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Enable different footers for odd and even pages.
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            // Set the page number style for the whole section to uppercase Roman numerals.
            doc.FirstSection.PageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;

            // Move the builder cursor to the even-page footer.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);

            // Insert the text and a PAGE field that will display the page number.
            builder.Write("Page ");
            builder.InsertField("PAGE", "");

            // Return to the main document body.
            builder.MoveToSection(0);

            // Add some content and page breaks to create multiple pages.
            builder.Writeln("First page content.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Second page content.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Third page content.");

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "EvenPageFooterRoman.docx");
            doc.Save(outputPath);
        }
    }
}
