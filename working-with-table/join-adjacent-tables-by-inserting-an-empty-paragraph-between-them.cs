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
        builder.Write("First table, cell 1");
        builder.EndRow();
        builder.EndTable();

        // Build the second table directly after the first (no paragraph between them).
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Second table, cell 1");
        builder.EndRow();
        builder.EndTable();

        // Locate the second table in the document.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        if (tables.Count < 2)
            throw new InvalidOperationException("The document does not contain two adjacent tables.");

        Table secondTable = (Table)tables[1];

        // Insert an empty paragraph before the second table to separate the tables.
        Paragraph emptyParagraph = new Paragraph(doc);
        secondTable.ParentNode.InsertBefore(emptyParagraph, secondTable);

        // Save the resulting document.
        string outputPath = "JoinedTables.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
