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

        // Create a new numbered list and reuse it for several paragraphs.
        List sharedList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Assign the list to the builder.
        builder.ListFormat.List = sharedList;

        // First level items.
        builder.ListFormat.ListLevelNumber = 0; // Level 0 = top level.
        builder.Writeln("First item");
        builder.Writeln("Second item");

        // Nested item (second level).
        builder.ListFormat.ListLevelNumber = 1; // Level 1 = nested level.
        builder.Writeln("Nested item");

        // Return to first level.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Third item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Add a normal paragraph after the list.
        builder.Writeln("Paragraph after the list.");

        // Validation: count paragraphs that are part of the list.
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
