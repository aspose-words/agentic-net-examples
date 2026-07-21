using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a builder to edit it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different headers/footers for odd and even pages.
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Header for odd-numbered pages (primary header).
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header for odd pages");

        // Header for even-numbered pages.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Write("Header for even pages");

        // Return to the main body of the first section.
        builder.MoveToSection(0);

        // Add a few pages of content to demonstrate the odd/even headers.
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Page {i}");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Prepare an output folder and file name.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "OddEvenHeaders.docx");

        // Save the document to disk.
        doc.Save(outputPath);
    }
}
