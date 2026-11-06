using System;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

namespace EvenPageFooterRoman
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Access the first (and only) section.
            Section section = doc.FirstSection;

            // Enable different footers for odd and even pages.
            section.PageSetup.OddAndEvenPagesHeaderFooter = true;

            // Add enough content to generate multiple pages.
            DocumentBuilder builder = new DocumentBuilder(doc);
            for (int i = 0; i < 30; i++)
            {
                builder.Writeln($"This is line {i + 1}.");
            }

            // Move the builder to the even-page footer.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);

            // Insert a PAGE field formatted with Roman numerals.
            // The field code "PAGE \\* ROMAN" tells Word to display the page number in Roman numerals.
            builder.InsertField("PAGE \\* ROMAN", null);

            // Save the document to disk.
            string outputPath = "EvenFooterRoman.docx";
            doc.Save(outputPath);
        }
    }
}
