using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    // Classes that match the JSON structure.
    public class TableData
    {
        public List<RowData> Rows { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
    }

    public class RowData
    {
        public List<CellData> Cells { get; set; }
    }

    public class CellData
    {
        public string Text { get; set; }
        public double Width { get; set; }                 // Width in points.
        public int ShadingColorArgb { get; set; }         // ARGB integer.
        public int VerticalAlignment { get; set; }        // 0=Top,1=Center,2=Bottom.
    }

    public static void Main()
    {
        // Sample JSON representing a table with formatting.
        string json = @"
        {
            ""Title"": ""Sample Table"",
            ""Description"": ""Table created from JSON data."",
            ""Rows"": [
                {
                    ""Cells"": [
                        { ""Text"": ""Header 1"", ""Width"": 150, ""ShadingColorArgb"": 0xFFE0E0E0, ""VerticalAlignment"": 1 },
                        { ""Text"": ""Header 2"", ""Width"": 150, ""ShadingColorArgb"": 0xFFE0E0E0, ""VerticalAlignment"": 1 }
                    ]
                },
                {
                    ""Cells"": [
                        { ""Text"": ""Row 1, Col 1"", ""Width"": 150, ""ShadingColorArgb"": 0xFFFFFFFF, ""VerticalAlignment"": 0 },
                        { ""Text"": ""Row 1, Col 2"", ""Width"": 150, ""ShadingColorArgb"": 0xFFFFFFFF, ""VerticalAlignment"": 0 }
                    ]
                },
                {
                    ""Cells"": [
                        { ""Text"": ""Row 2, Col 1"", ""Width"": 150, ""ShadingColorArgb"": 0xFFFFFFFF, ""VerticalAlignment"": 2 },
                        { ""Text"": ""Row 2, Col 2"", ""Width"": 150, ""ShadingColorArgb"": 0xFFFFFFFF, ""VerticalAlignment"": 2 }
                    ]
                }
            ]
        }";

        // Deserialize JSON into objects.
        TableData tableData = JsonConvert.DeserializeObject<TableData>(json);
        if (tableData == null || tableData.Rows == null)
            throw new InvalidOperationException("Failed to deserialize table data.");

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        Table table = builder.StartTable();

        // Build rows from JSON.
        for (int rowIndex = 0; rowIndex < tableData.Rows.Count; rowIndex++)
        {
            RowData rowData = tableData.Rows[rowIndex];

            foreach (CellData cellData in rowData.Cells)
            {
                // Insert a new cell.
                builder.InsertCell();

                // Apply cell formatting.
                builder.CellFormat.Width = cellData.Width;
                builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(cellData.ShadingColorArgb);
                builder.CellFormat.VerticalAlignment = (CellVerticalAlignment)cellData.VerticalAlignment;

                // Write the cell text.
                builder.Write(cellData.Text);
            }

            // End the current row.
            builder.EndRow();

            // After the first row exists we can safely set table title/description.
            if (rowIndex == 0)
            {
                if (!string.IsNullOrEmpty(tableData.Title))
                    table.Title = tableData.Title;
                if (!string.IsNullOrEmpty(tableData.Description))
                    table.Description = tableData.Description;
            }
        }

        // Finish the table.
        builder.EndTable();

        // Auto‑fit the table to its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DeserializedTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
