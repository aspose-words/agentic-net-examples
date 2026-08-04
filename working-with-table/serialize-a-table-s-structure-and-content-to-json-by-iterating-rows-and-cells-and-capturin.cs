using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using Newtonsoft.Json;

namespace AsposeWordsTableToJson
{
    // Simple DTOs for JSON serialization
    public class TableInfo
    {
        public string Alignment { get; set; }
        public double LeftIndent { get; set; }
        public List<RowInfo> Rows { get; set; } = new List<RowInfo>();
    }

    public class RowInfo
    {
        public double Height { get; set; }
        public string HeightRule { get; set; }
        public List<CellInfo> Cells { get; set; } = new List<CellInfo>();
    }

    public class CellInfo
    {
        public string Text { get; set; }
        public double Width { get; set; }
        public string VerticalAlignment { get; set; }
        public string Orientation { get; set; }
        public string ShadingBackgroundColor { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a sample table with formatting.
            Table table = builder.StartTable();

            // First row
            builder.InsertCell();
            builder.CellFormat.Width = 120;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            builder.Write("Header 1");

            builder.InsertCell();
            builder.CellFormat.Width = 150;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
            builder.Write("Header 2");
            builder.EndRow();

            // Second row
            builder.InsertCell();
            builder.RowFormat.Height = 30;
            builder.RowFormat.HeightRule = HeightRule.Exactly;
            builder.CellFormat.Width = 120;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Top;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
            builder.CellFormat.Orientation = TextOrientation.Upward;
            builder.Write("Row1 Col1");

            builder.InsertCell();
            builder.CellFormat.Width = 150;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Bottom;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
            builder.CellFormat.Orientation = TextOrientation.Downward;
            builder.Write("Row1 Col2");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Apply some table-level formatting.
            table.Alignment = TableAlignment.Center;
            table.LeftIndent = 20;

            // Save the document (required by the lifecycle rules).
            string docPath = "SampleTable.docx";
            doc.Save(docPath);

            // Traverse tables and collect structure + formatting.
            List<TableInfo> tablesInfo = new List<TableInfo>();
            NodeCollection tableNodes = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table tbl in tableNodes)
            {
                TableInfo tInfo = new TableInfo
                {
                    Alignment = tbl.Alignment.ToString(),
                    LeftIndent = tbl.LeftIndent
                };

                foreach (Row row in tbl.Rows)
                {
                    RowInfo rInfo = new RowInfo
                    {
                        Height = row.RowFormat.Height,
                        HeightRule = row.RowFormat.HeightRule.ToString()
                    };

                    foreach (Cell cell in row.Cells)
                    {
                        // Ensure the cell has at least one paragraph.
                        cell.EnsureMinimum();

                        CellInfo cInfo = new CellInfo
                        {
                            Text = cell.GetText().Trim(),
                            Width = cell.CellFormat.Width,
                            VerticalAlignment = cell.CellFormat.VerticalAlignment.ToString(),
                            Orientation = cell.CellFormat.Orientation.ToString(),
                            ShadingBackgroundColor = cell.CellFormat.Shading.BackgroundPatternColor.IsEmpty
                                ? null
                                : ColorTranslator.ToHtml(cell.CellFormat.Shading.BackgroundPatternColor)
                        };

                        rInfo.Cells.Add(cInfo);
                    }

                    tInfo.Rows.Add(rInfo);
                }

                tablesInfo.Add(tInfo);
            }

            // Serialize to JSON.
            string json = JsonConvert.SerializeObject(tablesInfo, Formatting.Indented);
            string jsonPath = "TableStructure.json";
            File.WriteAllText(jsonPath, json);

            // Output paths for verification (optional, not interactive).
            Console.WriteLine($"Document saved to: {Path.GetFullPath(docPath)}");
            Console.WriteLine($"JSON report saved to: {Path.GetFullPath(jsonPath)}");
        }
    }
}
