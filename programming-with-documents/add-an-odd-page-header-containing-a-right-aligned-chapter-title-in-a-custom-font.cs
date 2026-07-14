using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

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

        // Move the cursor to the odd‑page (primary) header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Set custom font for the chapter title.
        builder.Font.Name = "Times New Roman";
        builder.Font.Size = 14;
        builder.Font.Bold = true;

        // Align the paragraph to the right.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;

        // Write the chapter title into the odd‑page header.
        builder.Write("Chapter 1 – Introduction");

        // Return the cursor to the main document body.
        builder.MoveToSection(0);

        // Add some content and page breaks to create multiple pages.
        builder.Writeln("Content of page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 3.");

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OddPageHeader.docx");
        doc.Save(outputPath);
    }
}
