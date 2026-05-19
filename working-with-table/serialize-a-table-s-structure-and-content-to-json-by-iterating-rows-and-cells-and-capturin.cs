using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using Newtonsoft.Json;

namespace AsposeWordsTableJson
{
    // Classes that represent the serializable structure of a table.
    public class TableInfo
    {
        public string Alignment { get; set; }
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
        public int ShadingColorArgb { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x2 table with some formatting.
            Table table = builder.StartTable();

            // First row, first cell.
            builder.InsertCell();
            builder.CellFormat.Width = 100;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            builder.Write("Cell 1,1");

            // First row, second cell.
            builder.InsertCell();
            builder.CellFormat.Width = 150;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Bottom;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
            builder.Write("Cell 1,2");

            builder.EndRow();

            // Second row, first cell.
            builder.InsertCell();
            // Use default formatting for this cell.
            builder.Write("Cell 2,1");

            // Second row, second cell.
            builder.InsertCell();
            builder.Write("Cell 2,2");

            builder.EndTable();

            // Save the sample document.
            string docPath = Path.Combine(Environment.CurrentDirectory, "SampleTable.docx");
            doc.Save(docPath);

            // Extract table information.
            List<TableInfo> tablesInfo = new List<TableInfo>();
            NodeCollection tableNodes = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table tbl in tableNodes)
            {
                TableInfo tblInfo = new TableInfo
                {
                    Alignment = tbl.Alignment.ToString()
                };

                foreach (Row row in tbl.Rows)
                {
                    RowInfo rowInfo = new RowInfo
                    {
                        Height = row.RowFormat.Height,
                        HeightRule = row.RowFormat.HeightRule.ToString()
                    };

                    foreach (Cell cell in row.Cells)
                    {
                        CellInfo cellInfo = new CellInfo
                        {
                            Text = cell.ToString(SaveFormat.Text).Trim(),
                            Width = cell.CellFormat.Width,
                            VerticalAlignment = cell.CellFormat.VerticalAlignment.ToString(),
                            ShadingColorArgb = cell.CellFormat.Shading.BackgroundPatternColor.ToArgb()
                        };
                        rowInfo.Cells.Add(cellInfo);
                    }

                    tblInfo.Rows.Add(rowInfo);
                }

                tablesInfo.Add(tblInfo);
            }

            // Serialize the extracted information to JSON.
            string json = JsonConvert.SerializeObject(tablesInfo, Formatting.Indented);
            string jsonPath = Path.Combine(Environment.CurrentDirectory, "TableInfo.json");
            File.WriteAllText(jsonPath, json);

            // The program finishes here. No user interaction is required.
        }
    }
}
