using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First part of the line with Heading1 style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("This is a heading style. ");

        // Insert a style separator so the next text can have a different style on the same line.
        builder.InsertStyleSeparator();

        // Second part of the line with Quote style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Write("This is a quote style.");

        // Save the document to disk.
        string outputPath = "StyleSeparatorExample.docx";
        doc.Save(outputPath);

        // Search for paragraphs whose break is a style separator.
        int styleSeparatorCount = 0;
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (paragraph.BreakIsStyleSeparator)
                styleSeparatorCount++;
        }

        // Output the result.
        Console.WriteLine($"Found {styleSeparatorCount} paragraph(s) with a style separator.");
    }
}
