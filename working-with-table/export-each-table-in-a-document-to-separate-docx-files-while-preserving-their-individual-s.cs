using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class ExportTables
{
    public static void Main()
    {
        // Create a sample document with three tables, each having different styling.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First table – built‑in style.
        Table table1 = builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 1, Cell 1");
        builder.InsertCell();
        builder.Write("Table 1, Cell 2");
        builder.EndRow();
        builder.EndTable();
        table1.StyleIdentifier = StyleIdentifier.LightShadingAccent1;

        // Second table – custom style.
        TableStyle customStyle = (TableStyle)sourceDoc.Styles.Add(StyleType.Table, "MyCustomStyle");
        customStyle.Borders.Color = System.Drawing.Color.DarkGreen;
        customStyle.Borders.LineStyle = LineStyle.Single;
        customStyle.Shading.BackgroundPatternColor = System.Drawing.Color.LightYellow;

        Table table2 = builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 2, Cell 1");
        builder.InsertCell();
        builder.Write("Table 2, Cell 2");
        builder.EndRow();
        builder.EndTable();
        table2.Style = customStyle;

        // Third table – default formatting.
        Table table3 = builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 3, Cell 1");
        builder.InsertCell();
        builder.Write("Table 3, Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Ensure the document has a body to contain tables.
        sourceDoc.FirstSection.EnsureMinimum();

        // Convert any style‑based formatting to direct formatting so it is preserved when exported.
        sourceDoc.ExpandTableStylesToDirectFormatting();

        // Create output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Get all tables in the source document.
        NodeCollection tables = sourceDoc.GetChildNodes(NodeType.Table, true);

        for (int i = 0; i < tables.Count; i++)
        {
            Table srcTable = (Table)tables[i];

            // Create a new empty document for the exported table.
            Document destDoc = new Document();

            // Import the table node into the destination document, preserving its formatting.
            NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedNode = importer.ImportNode(srcTable, true);

            // Append the imported table to the body of the destination document.
            destDoc.FirstSection.Body.AppendChild(importedNode);

            // Save the individual table document.
            string outPath = Path.Combine(outputDir, $"Table_{i + 1}.docx");
            destDoc.Save(outPath);

            // Verify that the file was created.
            if (!File.Exists(outPath))
                throw new InvalidOperationException($"Failed to create output file: {outPath}");
        }

        Console.WriteLine($"Exported {tables.Count} tables to folder: {outputDir}");
    }
}
