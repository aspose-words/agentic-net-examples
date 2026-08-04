using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with 9 rows, each containing simple text.
        builder.StartTable();
        for (int i = 1; i <= 9; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Row {i}");
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table (the first and only table in the document).
        Table firstTable = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Create two empty tables that will hold the split parts.
        Table secondTable = new Table(doc);
        firstTable.ParentNode.InsertAfter(secondTable, firstTable);

        Table thirdTable = new Table(doc);
        secondTable.ParentNode.InsertAfter(thirdTable, secondTable);

        // ----- First split -----
        // Move rows 3‑8 (original indices) from firstTable to secondTable,
        // leaving rows 0‑2 in firstTable.
        for (int i = 0; i < 6; i++) // 6 rows to move (indices 3‑8)
        {
            Row rowToMove = firstTable.Rows[3]; // always the next row after the kept ones
            firstTable.Rows.RemoveAt(3);
            secondTable.Rows.Add(rowToMove);
        }

        // ----- Second split -----
        // Move rows 3‑5 (original rows 6‑8) from secondTable to thirdTable,
        // leaving rows 0‑2 in secondTable.
        for (int i = 0; i < 3; i++) // 3 rows to move
        {
            Row rowToMove = secondTable.Rows[3];
            secondTable.Rows.RemoveAt(3);
            thirdTable.Rows.Add(rowToMove);
        }

        // Save the document to verify the result.
        string outputPath = "SplitTable.docx";
        doc.Save(outputPath);

        // Output simple verification data.
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
        Console.WriteLine($"Table row counts -> First: {firstTable.Rows.Count}, Second: {secondTable.Rows.Count}, Third: {thirdTable.Rows.Count}");
    }
}
