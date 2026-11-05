using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    // Simple DTOs for JSON serialization
    public class CellInfo
    {
        public string Text { get; set; }
        public double Width { get; set; }
        public string ShadingColor { get; set; }
        public string ParagraphAlignment { get; set; }
    }

    public class RowInfo
    {
        public List<CellInfo> Cells { get; set; } = new List<CellInfo>();
    }

    public class TableInfo
    {
        public List<RowInfo> Rows { get; set; } = new List<RowInfo>();
    }

    public static void Main()
    {
        // Create a new document and builder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample table with formatting
        builder.StartTable();

        // First row, first cell
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        builder.CellFormat.Width = 100;
        builder.Write("Header 1");

        // First row, second cell
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
        builder.CellFormat.Width = 150;
        builder.Write("Header 2");
        builder.EndRow();

        // Second row, first cell
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
        builder.CellFormat.Width = 100;
        builder.Write("Row 1, Cell 1");

        // Second row, second cell
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
        builder.CellFormat.Width = 150;
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Save the document
        string docPath = "TableDoc.docx";
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new Exception("Document was not saved.");

        // Traverse tables and collect information
        List<TableInfo> tablesInfo = new List<TableInfo>();
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        foreach (Table table in tables)
        {
            TableInfo tInfo = new TableInfo();
            foreach (Row row in table.Rows)
            {
                RowInfo rInfo = new RowInfo();
                foreach (Cell cell in row.Cells)
                {
                    CellInfo cInfo = new CellInfo
                    {
                        Text = cell.GetText().TrimEnd('\a'), // Remove cell marker
                        Width = cell.CellFormat.Width,
                        ShadingColor = cell.CellFormat.Shading.BackgroundPatternColor.IsEmpty
                            ? "None"
                            : ColorTranslator.ToHtml(cell.CellFormat.Shading.BackgroundPatternColor),
                        ParagraphAlignment = cell.FirstParagraph?.ParagraphFormat?.Alignment.ToString() ?? "Unknown"
                    };
                    rInfo.Cells.Add(cInfo);
                }
                tInfo.Rows.Add(rInfo);
            }
            tablesInfo.Add(tInfo);
        }

        // Serialize to JSON
        string json = JsonConvert.SerializeObject(tablesInfo, Formatting.Indented);
        string jsonPath = "TableStructure.json";
        File.WriteAllText(jsonPath, json);
        if (!File.Exists(jsonPath))
            throw new Exception("JSON report was not created.");

        // Optionally, output paths (no interactive wait)
        Console.WriteLine($"Document saved to: {Path.GetFullPath(docPath)}");
        Console.WriteLine($"JSON report saved to: {Path.GetFullPath(jsonPath)}");
    }
}
