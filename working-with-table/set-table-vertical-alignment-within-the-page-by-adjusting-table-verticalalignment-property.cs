using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableVerticalAlignmentExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x2 table.
            builder.StartTable();

            builder.InsertCell();
            builder.Writeln("Cell 1");
            builder.InsertCell();
            builder.Writeln("Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("Cell 3");
            builder.InsertCell();
            builder.Writeln("Cell 4");
            builder.EndRow();

            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Aspose.Words does not expose a Table.VerticalAlignment property.
            // To approximate centering the table on the page vertically,
            // we can add empty paragraphs before and after the table.
            // Here we demonstrate a simple approach by inserting a paragraph
            // with a line break before the table, which can be adjusted as needed.
            // (If a future version provides Table.VerticalAlignment, replace this block with that property.)

            // Insert an empty paragraph before the table to push it down.
            Paragraph before = new Paragraph(doc);
            before.AppendChild(new Run(doc, "\n"));
            table.ParentNode.InsertBefore(before, table);

            // Insert an empty paragraph after the table to balance spacing.
            Paragraph after = new Paragraph(doc);
            after.AppendChild(new Run(doc, "\n"));
            table.ParentNode.InsertAfter(after, table);

            // Save the document.
            string outputPath = "TableVerticalAlignment.docx";
            doc.Save(outputPath);
        }
    }
}
