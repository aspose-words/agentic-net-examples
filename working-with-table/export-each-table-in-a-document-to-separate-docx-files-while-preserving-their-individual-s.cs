using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class ExportTables
{
    public static void Main()
    {
        // Create a source document containing two tables, each with its own style.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // ---------- First table with a custom style ----------
        TableStyle style1 = (TableStyle)sourceDoc.Styles.Add(StyleType.Table, "CustomStyle1");
        style1.Shading.BackgroundPatternColor = System.Drawing.Color.LightYellow;
        style1.Borders.Color = System.Drawing.Color.DarkBlue;
        style1.Borders.LineStyle = LineStyle.Single;
        style1.Borders.LineWidth = 1.5;

        Table table1 = builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Apply the custom style to the first table.
        table1.Style = style1;

        // ---------- Second table with default formatting ----------
        builder.Writeln(); // Paragraph break between tables.
        Table table2 = builder.StartTable();
        builder.InsertCell();
        builder.Write("A");
        builder.InsertCell();
        builder.Write("B");
        builder.InsertCell();
        builder.Write("C");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("1");
        builder.InsertCell();
        builder.Write("2");
        builder.InsertCell();
        builder.Write("3");
        builder.EndRow();

        builder.EndTable();

        // Convert any remaining style formatting to direct formatting.
        sourceDoc.ExpandTableStylesToDirectFormatting();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ExportedTables");
        Directory.CreateDirectory(outputDir);

        // Retrieve all tables from the source document.
        NodeCollection tables = sourceDoc.GetChildNodes(NodeType.Table, true);

        // Export each table to a separate DOCX file.
        for (int i = 0; i < tables.Count; i++)
        {
            Table srcTable = (Table)tables[i];

            // Create a new empty document to host the exported table.
            Document targetDoc = new Document();

            // Import the table into the target document, preserving its formatting.
            // Use the overload that takes the source document, not the table itself.
            NodeImporter importer = new NodeImporter(sourceDoc, targetDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedTable = importer.ImportNode(srcTable, true);

            // Append the imported table to the body of the target document.
            targetDoc.FirstSection.Body.AppendChild(importedTable);

            // Save the target document.
            string fileName = Path.Combine(outputDir, $"Table_{i + 1}.docx");
            targetDoc.Save(fileName, SaveFormat.Docx);

            // Verify that the file was created.
            if (!File.Exists(fileName))
                throw new InvalidOperationException($"Failed to create file: {fileName}");
        }
    }
}
