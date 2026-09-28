using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure the document uses different headers for odd and even pages.
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Move the builder to the primary (odd‑page) header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Align the paragraph to the right.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;

        // Set a custom font for the chapter title.
        builder.Font.Name = "Times New Roman";
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Font.Color = Color.DarkRed;

        // Write the chapter title.
        builder.Writeln("Chapter 1 – Introduction");

        // Return to the main document body.
        builder.MoveToDocumentEnd();

        // Add some body content to generate pages.
        for (int i = 1; i <= 30; i++)
        {
            builder.Writeln($"This is paragraph {i} of the document body.");
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
