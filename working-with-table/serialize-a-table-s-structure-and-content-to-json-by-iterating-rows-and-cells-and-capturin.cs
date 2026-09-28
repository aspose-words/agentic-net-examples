using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    // DTOs for JSON serialization
    public class TableInfo
    {
        public List<RowInfo> Rows { get; set; } = new List<RowInfo>();
    }

    public class RowInfo
    {
        public List<CellInfo> Cells { get; set; } = new List<CellInfo>();
    }

    public class CellInfo
    {
        public string Text { get; set; }
        public string ShadingColor { get; set; } // Hex string, null if no shading
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with formatting.
        builder.StartTable();

        // First row, first cell with shading.
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Header 1");
        builder.CurrentParagraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        // Apply shading to the current cell.
        Cell currentCell = builder.CurrentParagraph.ParentNode as Cell;
        if (currentCell != null)
            currentCell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;

        // First row, second cell without shading.
        builder.InsertCell();
        builder.Write("Header 2");
        // Ensure no shading.
        currentCell = builder.CurrentParagraph.ParentNode as Cell;
        if (currentCell != null)
            currentCell.CellFormat.Shading.BackgroundPatternColor = Color.Empty;

        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        currentCell = builder.CurrentParagraph.ParentNode as Cell;
        if (currentCell != null)
            currentCell.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        // No shading for this cell (default).

        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to verify creation.
        string docPath = "TableDoc.docx";
        doc.Save(docPath);

        // Retrieve the first table in the document.
        Table table = doc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (table == null)
            throw new InvalidOperationException("No table found in the document.");

        // Serialize table structure and content to JSON.
        TableInfo tableInfo = new TableInfo();

        foreach (Row row in table.Rows)
        {
            RowInfo rowInfo = new RowInfo();

            foreach (Cell cell in row.Cells)
            {
                CellInfo cellInfo = new CellInfo
                {
                    Text = cell.GetText().TrimEnd('\a')
                };

                Color shading = cell.CellFormat.Shading.BackgroundPatternColor;
                if (shading != Color.Empty && shading.A != 0)
                {
                    cellInfo.ShadingColor = ColorTranslator.ToHtml(shading);
                }
                else
                {
                    cellInfo.ShadingColor = null;
                }

                rowInfo.Cells.Add(cellInfo);
            }

            tableInfo.Rows.Add(rowInfo);
        }

        string json = JsonConvert.SerializeObject(tableInfo, Formatting.Indented);
        string jsonPath = "table.json";
        File.WriteAllText(jsonPath, json);

        // Validate that output files were created.
        if (!File.Exists(docPath))
            throw new FileNotFoundException($"Document file not found: {docPath}");
        if (!File.Exists(jsonPath))
            throw new FileNotFoundException($"JSON file not found: {jsonPath}");

        // Output paths (no interactive prompts).
        Console.WriteLine($"Document saved to: {Path.GetFullPath(docPath)}");
        Console.WriteLine($"Table JSON saved to: {Path.GetFullPath(jsonPath)}");
    }
}
