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

        // Create a bulleted list using a built‑in template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Assign the list to the builder.
        builder.ListFormat.List = bulletList;

        // Write top‑level list items.
        builder.Writeln("First bullet item");
        builder.Writeln("Second bullet item");

        // Switch to second level (nested list).
        builder.ListFormat.ListLevelNumber = 1; // Level index is zero‑based.
        builder.Writeln("Nested bullet item");

        // Return to top level.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Third bullet item");

        // End the first list.
        builder.ListFormat.RemoveNumbers();

        // Normal paragraph outside the list.
        builder.Writeln("This is a normal paragraph.");

        // Reuse the same list for another bullet.
        builder.ListFormat.List = bulletList;
        builder.Writeln("Another bullet after normal paragraph.");

        // End the second list.
        builder.ListFormat.RemoveNumbers();

        // Validation: count paragraphs that are part of a list.
        int listParagraphCount = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                listParagraphCount++;
        }

        // Output validation result.
        Console.WriteLine($"Number of list paragraphs: {listParagraphCount}");

        // Save the document.
        doc.Save("ListExample.docx");
    }
}
