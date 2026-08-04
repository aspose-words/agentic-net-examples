using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

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

            // Move to the odd-page (primary) header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Set custom font for the chapter title.
            builder.Font.Name = "Times New Roman";
            builder.Font.Size = 16;
            builder.Font.Bold = true;
            builder.Font.Color = System.Drawing.Color.DarkBlue;

            // Align the paragraph to the right.
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;

            // Write the chapter title.
            builder.Writeln("Chapter 1: Introduction");

            // Return to the main document body.
            builder.MoveToSection(0);

            // Add some sample pages to demonstrate the header.
            for (int i = 1; i <= 3; i++)
            {
                builder.Writeln($"This is the content of page {i}.");
                if (i < 3)
                {
                    builder.InsertBreak(BreakType.PageBreak);
                }
            }

            // Define output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OddPageHeader.docx");

            // Save the document.
            doc.Save(outputPath);
        }
    }
}
