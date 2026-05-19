using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string filePath = Path.Combine(outputDir, "ListNesting.docx");

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // Add 12 items. Levels 0‑8 (nine levels) are valid; levels 9+ will fall back to plain text.
        for (int i = 0; i < 12; i++)
        {
            builder.ListFormat.ListLevelNumber = i; // 0‑8 produce list formatting.
            builder.Writeln($"Level {i + 1}");
        }

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save(filePath);

        // Reload the document to verify list levels.
        Document loaded = new Document(filePath);
        int paragraphIndex = 0;

        foreach (Paragraph para in loaded.GetChildNodes(NodeType.Paragraph, true))
        {
            bool isListItem = para.ListFormat.IsListItem;
            bool expected = paragraphIndex < 9; // First nine paragraphs should be list items.
            Console.WriteLine($"Paragraph {paragraphIndex + 1}: IsListItem = {isListItem}, Expected = {expected}");
            paragraphIndex++;
        }
    }
}
