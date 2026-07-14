using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample source document that contains two tables,
        //    each with its own table style.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // ----- First table ------------------------------------------------
        Table table1 = builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Row1 Col1");
        builder.InsertCell();
        builder.Write("Row1 Col2");
        builder.EndRow();

        builder.EndTable();

        // Define a custom style for the first table.
        TableStyle style1 = (TableStyle)sourceDoc.Styles.Add(StyleType.Table, "MyStyle1");
        style1.RowStripe = 1;
        style1.CellSpacing = 2;
        style1.Shading.BackgroundPatternColor = Color.LightYellow;
        style1.Borders.Color = Color.Blue;
        style1.Borders.LineStyle = LineStyle.Single;
        table1.Style = style1;

        // ----- Second table -----------------------------------------------
        Table table2 = builder.StartTable();
        builder.InsertCell();
        builder.Write("A");
        builder.InsertCell();
        builder.Write("B");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("C");
        builder.InsertCell();
        builder.Write("D");
        builder.EndRow();

        builder.EndTable();

        // Define a custom style for the second table.
        TableStyle style2 = (TableStyle)sourceDoc.Styles.Add(StyleType.Table, "MyStyle2");
        style2.RowStripe = 2;
        style2.CellSpacing = 5;
        style2.Shading.BackgroundPatternColor = Color.LightGreen;
        style2.Borders.Color = Color.Red;
        style2.Borders.LineStyle = LineStyle.Double;
        table2.Style = style2;

        // Save the source document (optional, demonstrates loading later).
        string sourcePath = Path.Combine(artifactsDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document and extract each table.
        // -----------------------------------------------------------------
        Document doc = new Document(sourcePath);
        NodeCollection allTables = doc.GetChildNodes(NodeType.Table, true);

        for (int i = 0; i < allTables.Count; i++)
        {
            Table srcTable = (Table)allTables[i];

            // Create a new empty document for the current table.
            Document tableDoc = new Document();

            // Import the table node into the new document, preserving formatting.
            NodeImporter importer = new NodeImporter(srcTable.Document, tableDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedNode = importer.ImportNode(srcTable, true); // ImportNode takes only (Node, bool)
            Table importedTable = (Table)importedNode;

            // Append the imported table to the body of the new document.
            tableDoc.FirstSection.Body.AppendChild(importedTable);

            // Convert any table style to direct formatting so the visual appearance
            // stays the same when the style definition is not present in the new document.
            tableDoc.ExpandTableStylesToDirectFormatting();

            // Save the individual table document.
            string outPath = Path.Combine(artifactsDir, $"Table_{i + 1}.docx");
            tableDoc.Save(outPath);
        }
    }
}
