using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace HeaderFooterDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add content and configure headers.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Enable different headers for the first page and for odd/even pages.
            builder.PageSetup.DifferentFirstPageHeaderFooter = true;
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            // First page header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.Write("Header for the first page");

            // Even pages header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
            builder.Write("Header for even pages");

            // Odd pages (primary) header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Header for odd pages");

            // Return to the main body of the document.
            builder.MoveToSection(0);
            builder.Writeln("Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3");

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "HeadersAndFooters.docx");
            doc.Save(outputPath);
        }
    }
}
