using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure the document has at least one section.
        Section section = doc.FirstSection;

        // Create an odd-page header.
        HeaderFooter oddHeader = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        section.HeadersFooters.Add(oddHeader);

        // Build the header content: right‑aligned chapter title in a custom font.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
        builder.Font.Name = "Georgia";
        builder.Font.Size = 14;
        builder.Font.Bold = true;
        builder.Writeln("Chapter 1: Introduction");

        // Return builder to the main document body.
        builder.MoveToDocumentEnd();

        // Add some body text to generate multiple pages (optional).
        builder.Font.Name = "Times New Roman";
        builder.Font.Size = 12;
        for (int i = 0; i < 30; i++)
        {
            builder.Writeln($"This is line {i + 1} of the document body.");
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
