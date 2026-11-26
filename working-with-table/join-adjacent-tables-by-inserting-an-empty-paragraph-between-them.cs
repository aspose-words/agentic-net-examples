using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the first table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 1, Cell 1");
        builder.EndRow();
        builder.EndTable();

        // Build the second table directly after the first (no paragraph between them).
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 2, Cell 1");
        builder.EndRow();
        builder.EndTable();

        // Insert an empty paragraph between any adjacent tables.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        for (int i = 0; i < tables.Count - 1; i++)
        {
            Table firstTable = (Table)tables[i];
            Table secondTable = (Table)tables[i + 1];

            // Check if the tables are directly adjacent.
            if (firstTable.NextSibling == secondTable)
            {
                // Create an empty paragraph and insert it between the tables.
                Paragraph emptyParagraph = new Paragraph(doc);
                firstTable.ParentNode.InsertAfter(emptyParagraph, firstTable);
            }
        }

        // Save the resulting document.
        string outputPath = "JoinedTables.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
