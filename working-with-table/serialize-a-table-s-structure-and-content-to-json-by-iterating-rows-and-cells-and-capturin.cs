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
        public string Title { get; set; }
        public string Description { get; set; }
        public List<RowInfo> Rows { get; set; } = new List<RowInfo>();
    }

    public class RowInfo
    {
        public List<CellInfo> Cells { get; set; } = new List<CellInfo>();
    }

    public class CellInfo
    {
        public string Text { get; set; }
        public double Width { get; set; }
        public string VerticalAlignment { get; set; }
        public string ShadingColor { get; set; }
        public string Orientation { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a sample table.
            Table table = builder.StartTable();

            // First row – header cells.
            builder.InsertCell();
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            builder.CellFormat.Width = 120;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Second row – data cells.
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            // Third row – data cells.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Add some metadata to the table.
            table.Title = "Sample Table";
            table.Description = "Demonstrates table serialization to JSON.";
            table.Alignment = TableAlignment.Center;

            // Save the document to verify the table exists.
            const string docPath = "SampleTable.docx";
            doc.Save(docPath);

            // Verify the document was saved.
            if (!File.Exists(docPath))
                throw new Exception($"Failed to create the document '{docPath}'.");

            // Traverse all tables in the document and collect their data.
            List<TableInfo> tablesInfo = new List<TableInfo>();
            NodeCollection tableNodes = doc.GetChildNodes(NodeType.Table, true);

            foreach (Table tbl in tableNodes)
            {
                TableInfo tblInfo = new TableInfo
                {
                    Alignment = tbl.Alignment.ToString(),
                    Title = tbl.Title,
                    Description = tbl.Description
                };

                foreach (Row row in tbl.Rows)
                {
                    RowInfo rowInfo = new RowInfo();

                    foreach (Cell cell in row.Cells)
                    {
                        // Extract cell text.
                        string cellText = cell.ToString(SaveFormat.Text).Trim();

                        // Capture formatting details.
                        double width = cell.CellFormat.Width;
                        string verticalAlignment = cell.CellFormat.VerticalAlignment.ToString();
                        string orientation = cell.CellFormat.Orientation.ToString();

                        // Shading color may be empty; handle accordingly.
                        Color bgColor = cell.CellFormat.Shading.BackgroundPatternColor;
                        string shadingColor = bgColor.IsEmpty ? null : ColorTranslator.ToHtml(bgColor);

                        CellInfo cellInfo = new CellInfo
                        {
                            Text = cellText,
                            Width = width,
                            VerticalAlignment = verticalAlignment,
                            Orientation = orientation,
                            ShadingColor = shadingColor
                        };

                        rowInfo.Cells.Add(cellInfo);
                    }

                    tblInfo.Rows.Add(rowInfo);
                }

                tablesInfo.Add(tblInfo);
            }

            // Serialize the collected information to JSON.
            string json = JsonConvert.SerializeObject(tablesInfo, Formatting.Indented);
            const string jsonPath = "TableData.json";
            File.WriteAllText(jsonPath, json);

            // Verify the JSON file was created.
            if (!File.Exists(jsonPath))
                throw new Exception($"Failed to create the JSON report '{jsonPath}'.");

            // The program finishes here without waiting for user input.
        }
    }
}
