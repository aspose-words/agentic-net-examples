using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    // Model classes matching the JSON structure.
    public class TableData
    {
        public List<RowData> Rows { get; set; }
    }

    public class RowData
    {
        public List<CellData> Cells { get; set; }
    }

    public class CellData
    {
        public string Text { get; set; }
        public string BackgroundColor { get; set; } // HTML color string, e.g. "#D9E1F2"
        public bool Bold { get; set; }
    }

    public static void Main()
    {
        // Sample JSON representing a table with formatting.
        string json = @"
        {
            ""Rows"": [
                {
                    ""Cells"": [
                        { ""Text"": ""Header 1"", ""BackgroundColor"": ""#D9E1F2"", ""Bold"": true },
                        { ""Text"": ""Header 2"", ""BackgroundColor"": ""#D9E1F2"", ""Bold"": true }
                    ]
                },
                {
                    ""Cells"": [
                        { ""Text"": ""Row 1, Col 1"", ""BackgroundColor"": ""#FFFFFF"", ""Bold"": false },
                        { ""Text"": ""Row 1, Col 2"", ""BackgroundColor"": ""#FFFFFF"", ""Bold"": false }
                    ]
                },
                {
                    ""Cells"": [
                        { ""Text"": ""Row 2, Col 1"", ""BackgroundColor"": ""#FFFFFF"", ""Bold"": false },
                        { ""Text"": ""Row 2, Col 2"", ""BackgroundColor"": ""#FFFFFF"", ""Bold"": false }
                    ]
                }
            ]
        }";

        // Deserialize JSON into the model.
        TableData tableData = JsonConvert.DeserializeObject<TableData>(json);
        if (tableData == null || tableData.Rows == null)
            throw new InvalidOperationException("Failed to deserialize table data.");

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the table based on the deserialized data.
        builder.StartTable();

        foreach (RowData row in tableData.Rows)
        {
            if (row.Cells == null)
                continue;

            foreach (CellData cell in row.Cells)
            {
                // Insert a new cell.
                builder.InsertCell();

                // Apply background shading if a color is provided.
                if (!string.IsNullOrEmpty(cell.BackgroundColor))
                {
                    Color bgColor = ColorTranslator.FromHtml(cell.BackgroundColor);
                    builder.CellFormat.Shading.BackgroundPatternColor = bgColor;
                }

                // Apply bold formatting if required.
                builder.Font.Bold = cell.Bold;

                // Write the cell text.
                builder.Writeln(cell.Text);

                // Reset formatting for the next cell.
                builder.Font.Bold = false;
                builder.CellFormat.Shading.ClearFormatting();
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
