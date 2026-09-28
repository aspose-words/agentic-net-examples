using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the first table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Table 1 - Cell 1");
        builder.EndRow();
        builder.EndTable();

        // Build the second table directly after the first one (adjacent tables).
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Table 2 - Cell 1");
        builder.EndRow();
        builder.EndTable();

        // Insert an empty paragraph between any adjacent tables.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        for (int i = 0; i < tables.Count - 1; i++)
        {
            Table firstTable = (Table)tables[i];
            Table secondTable = (Table)tables[i + 1];

            // Ensure the tables are consecutive siblings.
            if (firstTable.NextSibling == secondTable)
            {
                // Create an empty paragraph.
                Paragraph emptyParagraph = new Paragraph(doc);
                // Insert the empty paragraph after the first table.
                firstTable.ParentNode.InsertAfter(emptyParagraph, firstTable);
            }
        }

        // Save the document.
        string outputPath = "JoinedTables.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // The program finishes automatically.
    }
}
