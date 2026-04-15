using System;
using Aspose.Words;

namespace AsposeWordsHeaderExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Enable different headers for odd and even pages.
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            // Move to the odd‑page header (HeaderPrimary) and add a right‑aligned chapter title.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.Font.Name = "Courier New";   // Custom font.
            builder.Font.Size = 14;
            builder.Write("Chapter 1");

            // Return to the main body of the document.
            builder.MoveToSection(0);

            // Add some content to generate multiple pages.
            for (int i = 1; i <= 5; i++)
            {
                builder.Writeln($"This is page {i} content.");
                if (i < 5)
                    builder.InsertBreak(BreakType.PageBreak);
            }

            // Save the document.
            const string outputPath = "OddPageHeader.docx";
            doc.Save(outputPath);
        }
    }
}
