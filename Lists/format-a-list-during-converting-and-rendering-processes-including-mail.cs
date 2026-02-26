using System;
using System.Data;
using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

class ListFormattingDemo
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // Create a numbered list based on a predefined template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Customize the first level of the list (e.g., font color and size).
        list.ListLevels[0].Font.Color = Color.DarkGreen;
        list.ListLevels[0].Font.Size = 12;
        // Use Arabic numbers for the first level.
        list.ListLevels[0].NumberStyle = NumberStyle.Arabic;
        list.ListLevels[0].StartAt = 1;

        // Use DocumentBuilder to add a simple list at the end of the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();
        builder.Writeln("List generated after loading document:");
        builder.ListFormat.List = list;
        for (int i = 0; i < 5; i++)
        {
            builder.Writeln($"Item {i + 1}");
        }
        builder.ListFormat.RemoveNumbers(); // End the list.

        // ------------------------------
        // Mail merge section.
        // ------------------------------

        // Prepare a DataTable with a single column "Product".
        DataTable table = new DataTable();
        table.Columns.Add("Product");
        table.Rows.Add("Apple");
        table.Rows.Add("Banana");
        table.Rows.Add("Cherry");

        // Insert merge fields into the document (one per line).
        builder.Writeln("\nMailMerge List:");
        foreach (DataRow _ in table.Rows)
        {
            builder.InsertField("MERGEFIELD Product \\* MERGEFORMAT");
            builder.Writeln();
        }

        // Execute the mail merge.
        doc.MailMerge.Execute(table);

        // Apply the list formatting to the paragraphs that now contain merged data.
        var mergedParagraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                                 .Cast<Paragraph>()
                                 .Where(p => !string.IsNullOrWhiteSpace(p.ToString(SaveFormat.Text)))
                                 .ToList();

        foreach (var para in mergedParagraphs)
        {
            para.ListFormat.List = list;
            // Use second level for distinction.
            para.ListFormat.ListLevelNumber = 1;
        }

        // ------------------------------
        // LINQ reporting section.
        // ------------------------------

        // Generate a sequence of numbers using LINQ and add them as paragraphs.
        var numbers = Enumerable.Range(1, 7);
        builder.Writeln("\nLINQ Generated List:");
        foreach (var n in numbers)
        {
            builder.Writeln($"Number {n}");
        }

        // Apply list formatting to the newly added LINQ paragraphs.
        var linqParagraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                                .Cast<Paragraph>()
                                .Where(p => p.ToString(SaveFormat.Text).StartsWith("Number"))
                                .ToList();

        foreach (var para in linqParagraphs)
        {
            para.ListFormat.List = list;
            // Use third level.
            para.ListFormat.ListLevelNumber = 2;
        }

        // NOTE: Aspose.Words for .NET Core/.NET 5+ does not expose a direct Print() method.
        // If printing is required, save the document to a printable format (e.g., PDF) and print
        // using external means. The following line is commented out to keep the example
        // compatible with all target frameworks.
        // doc.Print();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
