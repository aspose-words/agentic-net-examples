using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

namespace AsposeWordsTableToJson
{
    // Simple DTOs for JSON serialization
    public class TableInfo
    {
        public int RowCount { get; set; }
        public List<RowInfo> Rows { get; set; } = new List<RowInfo>();
    }

    public class RowInfo
    {
        public int CellCount { get; set; }
        public List<CellInfo> Cells { get; set; } = new List<CellInfo>();
    }

    public class CellInfo
    {
        public string Text { get; set; }
        public string ShadingColor { get; set; } // HTML hex, empty if none
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a sample document with a table
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start the table
            builder.StartTable();

            // Header row - apply shading to cells
            builder.InsertCell();
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            builder.Write("Header 1");

            builder.InsertCell();
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            builder.Write("Header 2");

            builder.EndRow();

            // Data row - no shading
            builder.InsertCell();
            builder.CellFormat.Shading.BackgroundPatternColor = Color.Empty;
            builder.Write("Cell 1");

            builder.InsertCell();
            builder.Write("Cell 2");

            builder.EndRow();

            // End the table
            builder.EndTable();

            // Save the sample document
            string docPath = "SampleTable.docx";
            doc.Save(docPath);
            if (!File.Exists(docPath))
                throw new Exception($"Failed to create document at {docPath}");

            // 2. Traverse tables and collect structure/content
            TableInfo tableInfo = new TableInfo();

            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table table in tables)
            {
                tableInfo.RowCount = table.Rows.Count;

                foreach (Row row in table.Rows)
                {
                    RowInfo rowInfo = new RowInfo
                    {
                        CellCount = row.Cells.Count
                    };

                    foreach (Cell cell in row.Cells)
                    {
                        // Capture cell text (trim end-of-cell markers)
                        string cellText = cell.GetText().TrimEnd('\a', '\r', '\n');

                        // Capture shading color if set
                        Color shading = cell.CellFormat.Shading.BackgroundPatternColor;
                        string shadingHex = shading.IsEmpty ? string.Empty : ColorTranslator.ToHtml(shading);

                        rowInfo.Cells.Add(new CellInfo
                        {
                            Text = cellText,
                            ShadingColor = shadingHex
                        });
                    }

                    tableInfo.Rows.Add(rowInfo);
                }
            }

            // 3. Serialize to JSON
            string json = JsonConvert.SerializeObject(tableInfo, Formatting.Indented);
            string jsonPath = "table.json";
            File.WriteAllText(jsonPath, json);

            // Verify JSON file creation
            if (!File.Exists(jsonPath))
                throw new Exception($"Failed to write JSON report at {jsonPath}");

            // Optional: output paths for confirmation
            Console.WriteLine($"Document saved to: {Path.GetFullPath(docPath)}");
            Console.WriteLine($"Table JSON saved to: {Path.GetFullPath(jsonPath)}");
        }
    }
}
