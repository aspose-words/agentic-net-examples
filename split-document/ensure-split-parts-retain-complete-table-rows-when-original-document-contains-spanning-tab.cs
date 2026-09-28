using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Create a sample document with a table that has 10 rows.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Build a table with two columns.
        builder.StartTable();
        for (int i = 1; i <= 10; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i} Cell 1");
            builder.InsertCell();
            builder.Write($"Row {i} Cell 2");
            builder.EndRow();
        }
        builder.EndTable();

        // Save the source document.
        string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // Split the table rows into parts of 5 rows each, preserving whole rows.
        int rowsPerPart = 5;
        List<Document> splitDocs = SplitTableByRows(sourceDoc, rowsPerPart);

        // Save split documents and validate their existence.
        for (int i = 0; i < splitDocs.Count; i++)
        {
            string partPath = $"Part{i + 1}.docx";
            splitDocs[i].Save(partPath);

            if (!File.Exists(partPath))
                throw new InvalidOperationException($"Failed to create split file: {partPath}");
        }

        // Program ends after successful execution.
    }

    private static List<Document> SplitTableByRows(Document sourceDoc, int rowsPerPart)
    {
        // Locate the first table in the source document.
        Table sourceTable = sourceDoc.GetChild(NodeType.Table, 0, true) as Table;
        if (sourceTable == null)
            throw new InvalidOperationException("Source document does not contain a table.");

        int totalRows = sourceTable.Rows.Count;
        int partCount = (totalRows + rowsPerPart - 1) / rowsPerPart;
        var result = new List<Document>();

        for (int partIndex = 0; partIndex < partCount; partIndex++)
        {
            int startRow = partIndex * rowsPerPart;
            int endRow = Math.Min(startRow + rowsPerPart, totalRows); // exclusive

            // Clone the whole table.
            Table clonedTable = (Table)sourceTable.Clone(true);

            // Remove rows that are not within the desired range.
            for (int rowIndex = clonedTable.Rows.Count - 1; rowIndex >= 0; rowIndex--)
            {
                if (rowIndex < startRow || rowIndex >= endRow)
                    clonedTable.Rows.RemoveAt(rowIndex);
            }

            // Create a new document and import the filtered table.
            Document partDoc = new Document();
            NodeImporter importer = new NodeImporter(sourceDoc, partDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedTable = importer.ImportNode(clonedTable, true);
            partDoc.FirstSection.Body.AppendChild(importedTable);

            result.Add(partDoc);
        }

        return result;
    }
}
