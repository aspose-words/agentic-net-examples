using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    // Represents a single cell in the JSON table definition.
    public class CellData
    {
        public string Text { get; set; }
        // Hex color string like "#FFCC00". Null means no shading.
        public string BackgroundColor { get; set; }
        // Width in points. Zero means default.
        public double Width { get; set; }
        // "Left", "Center", "Right", or "Justify". Null means default.
        public string HorizontalAlignment { get; set; }
    }

    // Represents a row in the JSON table definition.
    public class RowData
    {
        public List<CellData> Cells { get; set; }
    }

    // Represents the whole table in the JSON table definition.
    public class TableData
    {
        public List<RowData> Rows { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
    }

    static void Main()
    {
        // Sample JSON that stores table content and simple formatting.
        string json = @"
        {
            ""Title"": ""Sample Table"",
            ""Description"": ""Generated from JSON"",
            ""Rows"": [
                {
                    ""Cells"": [
                        { ""Text"": ""Header 1"", ""BackgroundColor"": ""#D9E1F2"", ""Width"": 150, ""HorizontalAlignment"": ""Center"" },
                        { ""Text"": ""Header 2"", ""BackgroundColor"": ""#D9E1F2"", ""Width"": 150, ""HorizontalAlignment"": ""Center"" }
                    ]
                },
                {
                    ""Cells"": [
                        { ""Text"": ""Row 1, Col 1"", ""BackgroundColor"": ""#FFFFFF"", ""Width"": 150, ""HorizontalAlignment"": ""Left"" },
                        { ""Text"": ""Row 1, Col 2"", ""BackgroundColor"": ""#FFFFFF"", ""Width"": 150, ""HorizontalAlignment"": ""Right"" }
                    ]
                },
                {
                    ""Cells"": [
                        { ""Text"": ""Row 2, Col 1"", ""BackgroundColor"": ""#FFFFFF"", ""Width"": 150, ""HorizontalAlignment"": ""Left"" },
                        { ""Text"": ""Row 2, Col 2"", ""BackgroundColor"": ""#FFFFFF"", ""Width"": 150, ""HorizontalAlignment"": ""Right"" }
                    ]
                }
            ]
        }";

        // Deserialize JSON into strongly‑typed objects.
        TableData tableData = JsonConvert.DeserializeObject<TableData>(json);

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building the table.
        Table table = builder.StartTable();

        // Iterate over rows.
        foreach (RowData row in tableData.Rows)
        {
            // Iterate over cells in the current row.
            foreach (CellData cell in row.Cells)
            {
                // Insert a new cell.
                builder.InsertCell();

                // Apply cell width if specified.
                if (cell.Width > 0)
                    builder.CellFormat.Width = cell.Width;

                // Apply background shading if a color is provided.
                if (!string.IsNullOrEmpty(cell.BackgroundColor))
                {
                    Color bg = ColorTranslator.FromHtml(cell.BackgroundColor);
                    builder.CellFormat.Shading.BackgroundPatternColor = bg;
                }

                // Set paragraph alignment based on the requested horizontal alignment.
                if (!string.IsNullOrEmpty(cell.HorizontalAlignment))
                {
                    switch (cell.HorizontalAlignment.Trim().ToLower())
                    {
                        case "left":
                            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                            break;
                        case "center":
                            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                            break;
                        case "right":
                            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                            break;
                        case "justify":
                            builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                            break;
                        default:
                            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                            break;
                    }
                }

                // Write the cell text.
                builder.Write(cell.Text ?? string.Empty);
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Apply stored table title and description.
        if (!string.IsNullOrEmpty(tableData.Title))
            table.Title = tableData.Title;
        if (!string.IsNullOrEmpty(tableData.Description))
            table.Description = tableData.Description;

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "TableFromJson.docx");
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output document.");

        // The program ends here without waiting for user input.
    }
}
