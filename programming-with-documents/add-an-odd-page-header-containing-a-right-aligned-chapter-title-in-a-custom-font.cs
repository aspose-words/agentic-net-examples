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

        // Enable different headers for odd and even pages.
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Move the cursor to the odd‑page (primary) header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Set the paragraph alignment to right and apply a custom font.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
        builder.Font.Name = "Courier New";
        builder.Font.Size = 14;
        builder.Font.Color = System.Drawing.Color.DarkBlue;

        // Write the chapter title that will appear on odd pages.
        builder.Write("Chapter 1: Introduction");

        // Return to the main document body.
        builder.MoveToSection(0);

        // Add some sample pages to demonstrate the header.
        builder.Writeln("Content of page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 3.");

        // Save the document.
        doc.Save("Output.docx");
    }
}
