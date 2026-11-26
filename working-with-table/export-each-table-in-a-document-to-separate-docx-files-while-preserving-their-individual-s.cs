using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample source document with two tables, each having a different style.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First table with "Table Grid" style.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 1 - Cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Table 1 - Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Apply style to the first table.
        Table firstTable = (Table)sourceDoc.GetChild(NodeType.Table, 0, true);
        firstTable.StyleIdentifier = StyleIdentifier.TableGrid;

        // Add a paragraph between tables.
        builder.Writeln();

        // Second table with "Light Shading Accent 1" style.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 2 - Cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Table 2 - Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Apply style to the second table.
        Table secondTable = (Table)sourceDoc.GetChild(NodeType.Table, 1, true);
        secondTable.StyleIdentifier = StyleIdentifier.LightShadingAccent1;

        // Ensure output directory exists.
        string outputDir = "ExportedTables";
        Directory.CreateDirectory(outputDir);

        // Iterate over each table in the source document.
        NodeCollection tables = sourceDoc.GetChildNodes(NodeType.Table, true);
        for (int i = 0; i < tables.Count; i++)
        {
            Table srcTable = (Table)tables[i];

            // Create a new empty document to hold the exported table.
            Document destDoc = new Document();

            // Import the table into the new document, preserving its formatting.
            NodeImporter importer = new NodeImporter(srcTable.Document, destDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedTable = importer.ImportNode(srcTable, true);

            // Append the imported table to the body of the destination document.
            destDoc.FirstSection.Body.AppendChild(importedTable);

            // Define the output file name.
            string outputPath = Path.Combine(outputDir, $"Table_{i + 1}.docx");

            // Save the document containing the single table.
            destDoc.Save(outputPath);

            // Validate that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create output file: {outputPath}");
        }

        // Optionally, save the source document for reference.
        string sourcePath = Path.Combine(outputDir, "SourceDocument.docx");
        sourceDoc.Save(sourcePath);
        if (!File.Exists(sourcePath))
            throw new InvalidOperationException($"Failed to save source document: {sourcePath}");
    }
}
