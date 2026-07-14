using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

namespace TableSerializationExample
{
    // Classes that represent the table structure for JSON serialization.
    public class TableInfo
    {
        public List<RowInfo> Rows { get; set; } = new List<RowInfo>();
    }

    public class RowInfo
    {
        public double Height { get; set; }
        public HeightRule HeightRule { get; set; }
        public List<CellInfo> Cells { get; set; } = new List<CellInfo>();
    }

    public class CellInfo
    {
        public string Text { get; set; }
        public double Width { get; set; }
        public CellVerticalAlignment VerticalAlignment { get; set; }
        public TextOrientation Orientation { get; set; }
        public Color? ShadingBackgroundColor { get; set; }
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

            // First row.
            builder.InsertCell();
            builder.CellFormat.Width = 100;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            builder.Write("Header 1");

            builder.InsertCell();
            builder.CellFormat.Width = 150;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
            builder.Write("Header 2");
            builder.EndRow();

            // Second row.
            builder.RowFormat.Height = 30;
            builder.RowFormat.HeightRule = HeightRule.Exactly;

            builder.InsertCell();
            builder.CellFormat.Width = 100;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Top;
            builder.CellFormat.Orientation = TextOrientation.Downward;
            builder.Write("Row1, Col1");

            builder.InsertCell();
            builder.CellFormat.Width = 150;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Bottom;
            builder.CellFormat.Orientation = TextOrientation.Upward;
            builder.Write("Row1, Col2");
            builder.EndRow();

            builder.EndTable();

            // Save the document to a local file.
            string docPath = "SampleTable.docx";
            doc.Save(docPath);

            // Load the document (optional, demonstrates loading).
            Document loadedDoc = new Document(docPath);

            // Extract table information.
            List<TableInfo> tablesInfo = new List<TableInfo>();
            NodeCollection tableNodes = loadedDoc.GetChildNodes(NodeType.Table, true);
            foreach (Table tbl in tableNodes)
            {
                TableInfo tblInfo = new TableInfo();

                foreach (Row row in tbl.Rows)
                {
                    RowInfo rowInfo = new RowInfo
                    {
                        Height = row.RowFormat.Height,
                        HeightRule = row.RowFormat.HeightRule
                    };

                    foreach (Cell cell in row.Cells)
                    {
                        // Ensure the cell has at least one paragraph to extract text.
                        cell.EnsureMinimum();

                        string cellText = cell.GetText().Trim('\a', '\r', '\n');

                        CellInfo cellInfo = new CellInfo
                        {
                            Text = cellText,
                            Width = cell.CellFormat.Width,
                            VerticalAlignment = cell.CellFormat.VerticalAlignment,
                            Orientation = cell.CellFormat.Orientation,
                            ShadingBackgroundColor = cell.CellFormat.Shading.BackgroundPatternColor.IsEmpty
                                ? (Color?)null
                                : cell.CellFormat.Shading.BackgroundPatternColor
                        };

                        rowInfo.Cells.Add(cellInfo);
                    }

                    tblInfo.Rows.Add(rowInfo);
                }

                tablesInfo.Add(tblInfo);
            }

            // Serialize the extracted information to JSON.
            string json = JsonConvert.SerializeObject(tablesInfo, Formatting.Indented);

            // Write JSON to a file.
            string jsonPath = "TableStructure.json";
            File.WriteAllText(jsonPath, json);

            // Indicate completion (no interactive input).
            Console.WriteLine("Table serialization completed.");
        }
    }
}
