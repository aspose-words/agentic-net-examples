using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list using a built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // Set the list level to three (third‑level item).
        builder.ListFormat.ListLevelNumber = 3;

        // Write the third‑level list item.
        builder.Writeln("Third level list item");

        // Remove list formatting after the list is complete.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        doc.Save("Output.docx");

        // Validation: count paragraphs that are part of a list.
        int listParagraphCount = doc.GetChildNodes(NodeType.Paragraph, true)
                                    .Cast<Paragraph>()
                                    .Count(p => p.ListFormat.IsListItem);

        Console.WriteLine($"Number of list paragraphs: {listParagraphCount}");
    }
}
