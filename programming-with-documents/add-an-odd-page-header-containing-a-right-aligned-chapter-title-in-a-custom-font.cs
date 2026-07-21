using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

namespace AsposeWordsHeaderExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Enable different headers/footers for odd and even pages.
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            // Move to the odd-page header (HeaderPrimary is used for odd pages when the above flag is true).
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Set paragraph alignment to right.
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;

            // Set custom font for the chapter title.
            builder.Font.Name = "Georgia";
            builder.Font.Size = 14;
            builder.Font.Bold = true;
            builder.Font.Color = System.Drawing.Color.DarkBlue;

            // Write the chapter title.
            builder.Writeln("Chapter 1: Introduction");

            // Return to the main document body.
            builder.MoveToSection(0);

            // Add enough content to generate multiple pages.
            for (int i = 1; i <= 3; i++)
            {
                builder.Writeln($"This is page {i} content.");
                builder.InsertBreak(BreakType.PageBreak);
            }

            // Prepare output folder.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Save the document.
            string outputPath = Path.Combine(artifactsDir, "OddPageHeader.docx");
            doc.Save(outputPath);
        }
    }
}
