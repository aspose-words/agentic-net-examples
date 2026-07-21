using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

namespace AsposeWordsTableJson
{
    // DTOs for JSON serialization
    public class TableInfo
    {
        public string Title { get; set; }
        public string Description { get; set; }
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
        public int ShadingColorArgb { get; set; }
        public double LeftPadding { get; set; }
        public double RightPadding { get; set; }
        public double TopPadding { get; set; }
        public double BottomPadding { get; set; }
        public string ParagraphAlignment { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a sample document with a table.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 3x2 table.
            Table table = builder.StartTable();

            // Header row
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // First data row
            builder.InsertCell();
            builder.Write("Row1 Col1");
            builder.InsertCell();
            builder.Write("Row1 Col2");
            builder.EndRow();

            // Second data row
            builder.InsertCell();
            builder.Write("Row2 Col1");
            builder.InsertCell();
            builder.Write("Row2 Col2");
            builder.EndRow();

            builder.EndTable();

            // Save the sample document.
            const string docPath = "SampleTable.docx";
            doc.Save(docPath);

            // 2. Traverse tables and collect structure + formatting.
            List<TableInfo> tablesInfo = new List<TableInfo>();

            NodeCollection tableNodes = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table tbl in tableNodes)
            {
                TableInfo ti = new TableInfo
                {
                    Title = tbl.Title,
                    Description = tbl.Description,
                    Alignment = tbl.Alignment.ToString(),
                    LeftIndent = tbl.LeftIndent
                };

                foreach (Row row in tbl.Rows)
                {
                    RowInfo ri = new RowInfo
                    {
                        Height = row.RowFormat.Height,
                        HeightRule = row.RowFormat.HeightRule.ToString()
                    };

                    foreach (Cell cell in row.Cells)
                    {
                        // Capture paragraph alignment of the first paragraph in the cell (if any)
                        string paraAlignment = cell.FirstParagraph?.ParagraphFormat?.Alignment.ToString() ?? "Left";

                        CellInfo ci = new CellInfo
                        {
                            Text = cell.ToString(SaveFormat.Text).Trim(),
                            Width = cell.CellFormat.Width,
                            VerticalAlignment = cell.CellFormat.VerticalAlignment.ToString(),
                            Orientation = cell.CellFormat.Orientation.ToString(),
                            ShadingColorArgb = cell.CellFormat.Shading.BackgroundPatternColor.ToArgb(),
                            LeftPadding = cell.CellFormat.LeftPadding,
                            RightPadding = cell.CellFormat.RightPadding,
                            TopPadding = cell.CellFormat.TopPadding,
                            BottomPadding = cell.CellFormat.BottomPadding,
                            ParagraphAlignment = paraAlignment
                        };

                        ri.Cells.Add(ci);
                    }

                    ti.Rows.Add(ri);
                }

                tablesInfo.Add(ti);
            }

            // 3. Serialize the collected information to JSON and save it.
            string json = JsonConvert.SerializeObject(tablesInfo, Formatting.Indented);
            const string jsonPath = "TablesInfo.json";
            File.WriteAllText(jsonPath, json);

            // Optional: output paths for verification.
            Console.WriteLine($"Document saved to: {Path.GetFullPath(docPath)}");
            Console.WriteLine($"Table metadata JSON saved to: {Path.GetFullPath(jsonPath)}");
        }
    }
}
