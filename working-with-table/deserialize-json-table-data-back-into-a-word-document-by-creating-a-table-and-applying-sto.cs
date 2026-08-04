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
    public class TableJson
    {
        public List<RowJson> Rows { get; set; }
    }

    public class RowJson
    {
        public List<CellJson> Cells { get; set; }
    }

    public class CellJson
    {
        public string Text { get; set; }
        // Optional background color in HTML hex format, e.g. "#FFCC00".
        public string BackgroundColor { get; set; }
    }

    public static void Main()
    {
        // Sample JSON representing a table with formatting.
        string json = @"
        {
            ""Rows"": [
                {
                    ""Cells"": [
                        { ""Text"": ""Item"", ""BackgroundColor"": ""#D9E1F2"" },
                        { ""Text"": ""Quantity"", ""BackgroundColor"": ""#D9E1F2"" }
                    ]
                },
                {
                    ""Cells"": [
                        { ""Text"": ""Apples"", ""BackgroundColor"": ""#FFFFFF"" },
                        { ""Text"": ""10"", ""BackgroundColor"": ""#FFFFFF"" }
                    ]
                },
                {
                    ""Cells"": [
                        { ""Text"": ""Bananas"", ""BackgroundColor"": ""#FFFFFF"" },
                        { ""Text"": ""20"", ""BackgroundColor"": ""#FFFFFF"" }
                    ]
                }
            ]
        }";

        // Deserialize the JSON into objects.
        TableJson tableData = JsonConvert.DeserializeObject<TableJson>(json);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building the table.
        Table table = builder.StartTable();

        // Iterate over rows and cells, applying stored formatting.
        foreach (RowJson rowJson in tableData.Rows)
        {
            foreach (CellJson cellJson in rowJson.Cells)
            {
                // Insert a new cell.
                builder.InsertCell();

                // Reset any previous cell formatting.
                builder.CellFormat.ClearFormatting();

                // Apply background shading if a color is provided.
                if (!string.IsNullOrEmpty(cellJson.BackgroundColor))
                {
                    Color bg = ColorTranslator.FromHtml(cellJson.BackgroundColor);
                    builder.CellFormat.Shading.BackgroundPatternColor = bg;
                }

                // Write the cell text.
                builder.Write(cellJson.Text ?? string.Empty);
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DeserializedTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("Failed to create the output Word document.");
        }
    }
}
