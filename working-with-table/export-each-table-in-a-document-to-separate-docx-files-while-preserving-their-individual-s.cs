using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace ExportTablesExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample source document containing multiple tables with distinct styles.
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            // First table with a built‑in style.
            Table table1 = builder.StartTable();
            builder.InsertCell();
            builder.Write("Table 1, Cell 1");
            builder.InsertCell();
            builder.Write("Table 1, Cell 2");
            builder.EndRow();
            builder.EndTable();
            table1.StyleIdentifier = StyleIdentifier.LightShadingAccent1;

            // Second table with a custom style.
            Table table2 = builder.StartTable();
            builder.InsertCell();
            builder.Write("Table 2, Cell 1");
            builder.InsertCell();
            builder.Write("Table 2, Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Create a custom table style.
            TableStyle customStyle = (TableStyle)sourceDoc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
            customStyle.Shading.BackgroundPatternColor = System.Drawing.Color.LightYellow;
            customStyle.Borders.Color = System.Drawing.Color.DarkBlue;
            customStyle.Borders.LineStyle = LineStyle.Single;
            customStyle.Borders.LineWidth = 1.5;
            table2.Style = customStyle;

            // Convert any style‑based formatting to direct formatting so that it is retained after export.
            sourceDoc.ExpandTableStylesToDirectFormatting();

            // Retrieve all tables in the source document.
            NodeCollection tables = sourceDoc.GetChildNodes(NodeType.Table, true);

            // Export each table to its own DOCX file.
            for (int i = 0; i < tables.Count; i++)
            {
                Table srcTable = (Table)tables[i];

                // Create a new blank document for the exported table.
                Document destDoc = new Document();

                // Import the table node into the destination document, preserving formatting.
                NodeImporter importer = new NodeImporter(srcTable.Document, destDoc, ImportFormatMode.KeepSourceFormatting);
                Node importedTable = importer.ImportNode(srcTable, true);

                // Append the imported table to the body of the destination document.
                destDoc.FirstSection.Body.AppendChild(importedTable);

                // Define the output file name.
                string outFileName = $"Table_{i + 1}.docx";

                // Save the document.
                destDoc.Save(outFileName);

                // Verify that the file was created.
                if (!File.Exists(outFileName))
                    throw new InvalidOperationException($"Failed to create output file: {outFileName}");
            }
        }
    }
}
