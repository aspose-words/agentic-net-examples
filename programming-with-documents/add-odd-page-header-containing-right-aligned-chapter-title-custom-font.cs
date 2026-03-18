using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable separate headers for odd and even pages.
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // ----- Odd‑page header (primary header) -----
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Align the paragraph to the right.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;

        // Apply a custom font to the header text.
        builder.Font.Name = "Calibri";      // replace with your custom font name
        builder.Font.Size = 14;
        builder.Font.Bold = true;

        // Write the chapter title.
        builder.Writeln("Chapter 1 – Introduction");

        // Return the cursor to the main document body.
        builder.MoveToSection(0);

        // Add some body text with page breaks to demonstrate odd/even pages.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"This is paragraph {i} on page {i}.");
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document.
        doc.Save("OddPageHeader.docx");
    }
}
