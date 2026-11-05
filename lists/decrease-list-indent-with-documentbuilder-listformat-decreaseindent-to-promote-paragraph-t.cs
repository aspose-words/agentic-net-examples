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

        // Create a numbered list and assign it to the builder.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;

        // First level item.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Item 1");

        // Second level (nested) item.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Item 1.1");

        // Third level item.
        builder.ListFormat.ListLevelNumber = 2;
        builder.Writeln("Item 1.1.1");

        // Promote the last item to a higher level by decreasing the level number.
        builder.ListFormat.ListLevelNumber = 1; // Moves from level 2 to level 1.
        builder.Writeln("Item 1.2 (promoted)");

        // Return to top level for the next item.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Item 2");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Validation: count paragraphs that are part of the list.
        int listParagraphs = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                listParagraphs++;
        }

        // Output validation result.
        builder.Writeln($"List paragraphs count: {listParagraphs}");

        // Save the document to disk.
        doc.Save("ListIndentExample.docx");
    }
}
