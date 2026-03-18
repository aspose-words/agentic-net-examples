using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

namespace TableToJsonExample
{
    // Simple DTOs that will be serialized to JSON.
    public class TableInfo
    {
        public int Index { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public List<RowInfo> Rows { get; set; } = new List<RowInfo>();
    }

    public class RowInfo
    {
        public int Index { get; set; }
        public List<CellInfo> Cells { get; set; } = new List<CellInfo>();
    }

    public class CellInfo
    {
        public int Index { get; set; }
        public string Text { get; set; }
        public double Width { get; set; }                     // Width in points.
        public string ShadingBackgroundColor { get; set; }   // Hex string.
        public BorderInfo[] Borders { get; set; } = new BorderInfo[6]; // Left, Right, Top, Bottom, DiagonalDown, DiagonalUp.
    }

    public class BorderInfo
    {
        public string Type { get; set; }          // BorderType name.
        public string LineStyle { get; set; }     // LineStyle name.
        public double Width { get; set; }         // Width in points.
        public string Color { get; set; }         // Hex string.
    }

    class Program
    {
        static void Main()
        {
            // Load an existing Word document that contains tables.
            // (Lifecycle rule: use the provided load constructor.)
            Document doc = new Document("Input.docx");

            // Prepare a list to hold information about all tables.
            List<TableInfo> tablesInfo = new List<TableInfo>();

            // Iterate through each table in the first section's body.
            TableCollection tables = doc.FirstSection.Body.Tables;
            for (int t = 0; t < tables.Count; t++)
            {
                Table table = tables[t];
                TableInfo tInfo = new TableInfo
                {
                    Index = t,
                    RowCount = table.Rows.Count,
                    ColumnCount = table.FirstRow?.Count ?? 0
                };

                // Iterate rows.
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    Row row = table.Rows[r];
                    RowInfo rInfo = new RowInfo { Index = r };

                    // Iterate cells.
                    for (int c = 0; c < row.Cells.Count; c++)
                    {
                        Cell cell = row.Cells[c];
                        CellInfo cInfo = new CellInfo
                        {
                            Index = c,
                            // Get plain text of the cell, trimming trailing control characters.
                            Text = cell.ToString(SaveFormat.Text).Trim(),
                            Width = cell.CellFormat.Width,
                            ShadingBackgroundColor = ColorToHex(cell.CellFormat.Shading.BackgroundPatternColor)
                        };

                        // Capture borders (six possible types).
                        foreach (BorderType bt in Enum.GetValues(typeof(BorderType)))
                        {
                            var border = cell.CellFormat.Borders[bt];
                            cInfo.Borders[(int)bt] = new BorderInfo
                            {
                                Type = bt.ToString(),
                                LineStyle = border.LineStyle.ToString(),
                                Width = border.LineWidth,
                                Color = ColorToHex(border.Color)
                            };
                        }

                        rInfo.Cells.Add(cInfo);
                    }

                    tInfo.Rows.Add(rInfo);
                }

                tablesInfo.Add(tInfo);
            }

            // Serialize the collected structure to JSON.
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(tablesInfo, jsonOptions);

            // Save JSON to a file.
            File.WriteAllText("TablesStructure.json", json);

            // (Lifecycle rule: optionally save the document if any modifications were made.)
            // doc.Save("Output.docx");
        }

        // Helper to convert System.Drawing.Color to hex string; handles empty colors.
        private static string ColorToHex(System.Drawing.Color color)
        {
            if (color.IsEmpty)
                return null;
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }
    }
}
