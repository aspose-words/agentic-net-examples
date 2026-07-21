using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing two tables with different built‑in styles.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First table.
        Table table1 = builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 1, Cell 1");
        builder.InsertCell();
        builder.Write("Table 1, Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Table 1, Cell 3");
        builder.InsertCell();
        builder.Write("Table 1, Cell 4");
        builder.EndRow();
        builder.EndTable();
        table1.StyleIdentifier = StyleIdentifier.LightShadingAccent1;

        // Second table.
        Table table2 = builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 2, Cell 1");
        builder.InsertCell();
        builder.Write("Table 2, Cell 2");
        builder.EndRow();
        builder.EndTable();
        table2.StyleIdentifier = StyleIdentifier.MediumShading1Accent2;

        // Convert style formatting to direct formatting so it is retained after export.
        sourceDoc.ExpandTableStylesToDirectFormatting();

        // Prepare the output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputTables");
        Directory.CreateDirectory(outputDir);

        // Get all tables from the source document.
        NodeCollection tables = sourceDoc.GetChildNodes(NodeType.Table, true);

        for (int i = 0; i < tables.Count; i++)
        {
            Table srcTable = (Table)tables[i];

            // Create a new empty document for the individual table.
            Document tableDoc = new Document();

            // Import the table node from the source document into the new document.
            // Use the source document (not the table) as the first argument of NodeImporter.
            NodeImporter importer = new NodeImporter(sourceDoc, tableDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedTable = importer.ImportNode(srcTable, true);

            // Append the imported table to the body of the new document.
            tableDoc.FirstSection.Body.AppendChild(importedTable);

            // Save the individual table document.
            string outPath = Path.Combine(outputDir, $"Table_{i + 1}.docx");
            tableDoc.Save(outPath);

            // Verify that the file was created.
            if (!File.Exists(outPath))
                throw new InvalidOperationException($"Failed to create output file: {outPath}");
        }
    }
}
