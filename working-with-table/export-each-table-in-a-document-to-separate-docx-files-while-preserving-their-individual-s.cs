using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class ExportTablesExample
{
    public static void Main()
    {
        // Create a sample document containing multiple tables, each with a distinct style.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Document containing multiple tables with individual styles.");

        // First table – apply Light List Accent 1 style.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 1 - Cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Table 1 - Cell 2");
        builder.EndRow();
        builder.EndTable();

        Table firstTable = sourceDoc.LastSection.Body.Tables[sourceDoc.LastSection.Body.Tables.Count - 1];
        firstTable.StyleIdentifier = StyleIdentifier.LightListAccent1;

        // Second table – apply Light Grid Accent 2 style.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 2 - Cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Table 2 - Cell 2");
        builder.EndRow();
        builder.EndTable();

        Table secondTable = sourceDoc.LastSection.Body.Tables[sourceDoc.LastSection.Body.Tables.Count - 1];
        secondTable.StyleIdentifier = StyleIdentifier.LightGridAccent2;

        // Gather all tables from the source document.
        List<Table> tables = new List<Table>();
        foreach (Node node in sourceDoc.GetChildNodes(NodeType.Table, true))
        {
            tables.Add((Table)node);
        }

        // Export each table to its own DOCX file while preserving its style.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Create a new document that will contain the single table.
            Document tableDoc = new Document();

            // Use the NodeImporter that takes source and destination documents.
            NodeImporter importer = new NodeImporter(sourceDoc, tableDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedTable = importer.ImportNode(table, true);

            // Append the imported table to the body of the new document.
            tableDoc.FirstSection.Body.AppendChild(importedTable);

            // Define and save the output file.
            string outputFileName = $"Table_{i + 1}.docx";
            tableDoc.Save(outputFileName);

            // Verify that the file was created.
            if (!File.Exists(outputFileName))
                throw new Exception($"Failed to create output file: {outputFileName}");
        }
    }
}
