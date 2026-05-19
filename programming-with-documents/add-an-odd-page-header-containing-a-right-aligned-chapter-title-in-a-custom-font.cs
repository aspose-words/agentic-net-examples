using System;
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
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Enable different headers for odd and even pages.
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            // Move to the primary header (used for odd pages).
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Set custom font for the chapter title.
            builder.Font.Name = "Georgia";
            builder.Font.Size = 16;
            builder.Font.Bold = true;

            // Align the paragraph to the right.
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;

            // Write the chapter title.
            builder.Writeln("Chapter 1 – Introduction");

            // Return to the main document body.
            builder.MoveToSection(0);

            // Add some pages to demonstrate the header.
            builder.Writeln("Content of page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Content of page 2.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Content of page 3.");

            // Save the document.
            doc.Save("OddPageHeader.docx");
        }
    }
}
