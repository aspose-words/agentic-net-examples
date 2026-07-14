using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

public class ListNestingExample
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list that supports up to 9 levels (0‑8).
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Add items for levels 0 through 10 (11 items). Levels 0‑8 should be formatted as a list,
        // levels 9 and 10 will fall back to plain text because the list supports only nine levels.
        for (int i = 0; i <= 10; i++)
        {
            builder.ListFormat.ListLevelNumber = i; // 0‑8 are valid, 9+ exceed the limit.
            builder.Writeln($"Level {i}");
        }

        // End the list formatting.
        builder.ListFormat.List = null;

        // Save the document.
        string outPath = Path.Combine(artifactsDir, "ListNesting.docx");
        doc.Save(outPath);

        // Reload the document to verify the paragraph properties.
        Document loaded = new Document(outPath);
        NodeCollection paragraphs = loaded.GetChildNodes(NodeType.Paragraph, true);

        for (int i = 0; i < paragraphs.Count; i++)
        {
            Paragraph para = (Paragraph)paragraphs[i];
            // Skip empty paragraphs that may be added automatically.
            string text = para.GetText().Trim();
            if (string.IsNullOrEmpty(text))
                continue;

            bool isListItem = para.ListFormat.IsListItem;
            int listLevel = para.ListFormat.ListLevelNumber; // 0‑8 for list items, 0 for plain text.

            Console.WriteLine($"Paragraph {i}: Text=\"{text}\", IsListItem={isListItem}, ListLevel={listLevel}");
        }
    }
}
