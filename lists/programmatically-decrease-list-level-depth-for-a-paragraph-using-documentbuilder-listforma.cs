using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list using a built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // First level item.
        builder.ListFormat.ListLevelNumber = 0; // top level
        builder.Writeln("Item 1");

        // Increase indent to create a nested level.
        builder.ListFormat.ListLevelNumber = 1; // second level
        builder.Writeln("Nested Item 2");

        // Conditional decrease of indent.
        bool shouldDecrease = true; // Replace with actual condition as needed.
        if (shouldDecrease)
        {
            // Decrease indent brings the next paragraph back to the parent level.
            builder.ListFormat.ListLevelNumber = 0; // back to top level
        }
        else
        {
            // Keep the current level if not decreasing.
            builder.ListFormat.ListLevelNumber = 1;
        }

        // Back to the original level (or adjusted level based on condition).
        builder.Writeln("Item 3");

        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // Simple validation: count paragraphs that are part of a list.
        int listParagraphs = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                listParagraphs++;
        }

        Console.WriteLine($"Document contains {listParagraphs} list paragraphs.");

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
