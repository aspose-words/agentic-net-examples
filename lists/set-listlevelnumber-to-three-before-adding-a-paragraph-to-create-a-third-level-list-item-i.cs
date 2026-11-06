using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a bulleted list and assign it to the builder.
        List list = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = list;

        // First‑level list item.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("First level item");

        // Second‑level list item.
        builder.ListFormat.ListLevelNumber = 2;
        builder.Writeln("Second level item");

        // Third‑level list item (required for the task).
        builder.ListFormat.ListLevelNumber = 3;
        builder.Writeln("Third level item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        doc.Save("NestedList.docx");

        // Validation: count paragraphs that are part of a list.
        int listParagraphCount = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                listParagraphCount++;
        }

        Console.WriteLine($"Document contains {listParagraphCount} list items.");
    }
}
