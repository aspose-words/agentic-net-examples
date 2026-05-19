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

        // ---------- First table ----------
        Table table1 = builder.StartTable();
        builder.InsertCell();
        builder.Write("Table1 Cell1");
        builder.InsertCell();
        builder.Write("Table1 Cell2");
        builder.EndRow();
        builder.EndTable();

        // Insert an empty paragraph between the tables.
        Paragraph emptyParagraph = builder.InsertParagraph(); // No text added, remains empty.

        // ---------- Second table ----------
        Table table2 = builder.StartTable();
        builder.InsertCell();
        builder.Write("Table2 Cell1");
        builder.InsertCell();
        builder.Write("Table2 Cell2");
        builder.EndRow();
        builder.EndTable();

        // ----- Validation -----
        // Ensure exactly two tables exist.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        if (tables.Count != 2)
            throw new Exception($"Expected 2 tables, but found {tables.Count}.");

        // Verify that a paragraph node follows the first table.
        Node firstTable = tables[0];
        Node nodeAfterFirstTable = firstTable.NextSibling;
        if (nodeAfterFirstTable == null || nodeAfterFirstTable.NodeType != NodeType.Paragraph)
            throw new Exception("The node after the first table is not a paragraph.");

        Paragraph betweenParagraph = (Paragraph)nodeAfterFirstTable;
        // The paragraph should be empty (no visible text).
        if (!string.IsNullOrWhiteSpace(betweenParagraph.GetText()))
            throw new Exception("The paragraph between tables is not empty.");

        // Save the document.
        string outputPath = "JoinedTables.docx";
        doc.Save(outputPath);

        // Confirm the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Indicate successful execution.
        Console.WriteLine("Document with joined tables created successfully.");
    }
}
