using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure the first (and only) section to have different odd/even headers.
        builder.PageSetup.DifferentFirstPageHeaderFooter = false;
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Move to the odd-page (primary) header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Set paragraph alignment to right.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;

        // Set custom font for the chapter title.
        builder.Font.Name = "Times New Roman";
        builder.Font.Size = 16;
        builder.Font.Bold = true;

        // Write the chapter title.
        builder.Writeln("Chapter 1: Introduction");

        // Return to the main document body.
        builder.MoveToDocumentEnd();

        // Add some sample content to generate multiple pages (optional).
        for (int i = 0; i < 30; i++)
        {
            builder.Writeln($"This is line {i + 1} of the document body.");
        }

        // Save the document.
        string outputPath = "OddPageHeader.docx";
        doc.Save(outputPath);
    }
}
