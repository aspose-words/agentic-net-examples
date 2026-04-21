using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    // Classes that represent the JSON structure.
    public class TableData
    {
        public string Title { get; set; }
        public string Description { get; set; }
        public RowData[] Rows { get; set; }
    }

    public class RowData
    {
        public CellData[] Cells { get; set; }
    }

    public class CellData
    {
        public string Text { get; set; }
        // Optional background color in HTML hex format, e.g. "#FFCCCC".
        public string BackgroundColor { get; set; }
    }

    public static void Main()
    {
        // Sample JSON that describes a table with formatting.
        string json = @"
        {
            ""Title"": ""Sample Table"",
            ""Description"": ""Table created from JSON data"",
            ""Rows"": [
                {
                    ""Cells"": [
                        { ""Text"": ""Item"", ""BackgroundColor"": ""#FFCCCC"" },
                        { ""Text"": ""Quantity"", ""BackgroundColor"": ""#FFCCCC"" }
                    ]
                },
                {
                    ""Cells"": [
                        { ""Text"": ""Apples"", ""BackgroundColor"": ""#FFFFFF"" },
                        { ""Text"": ""20"", ""BackgroundColor"": ""#FFFFFF"" }
                    ]
                },
                {
                    ""Cells"": [
                        { ""Text"": ""Bananas"", ""BackgroundColor"": ""#FFFFFF"" },
                        { ""Text"": ""40"", ""BackgroundColor"": ""#FFFFFF"" }
                    ]
                }
            ]
        }";

        // Deserialize the JSON into objects.
        TableData tableData = JsonConvert.DeserializeObject<TableData>(json);

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building the table.
        Table table = builder.StartTable();

        // Iterate over rows and cells, applying cell-level formatting.
        foreach (RowData rowData in tableData.Rows)
        {
            foreach (CellData cellData in rowData.Cells)
            {
                // Insert a new cell.
                builder.InsertCell();

                // Apply background shading if a color is provided.
                if (!string.IsNullOrEmpty(cellData.BackgroundColor))
                {
                    Color bgColor = ColorTranslator.FromHtml(cellData.BackgroundColor);
                    builder.CellFormat.Shading.BackgroundPatternColor = bgColor;
                }
                else
                {
                    // Ensure no shading is applied when color is absent.
                    builder.CellFormat.Shading.BackgroundPatternColor = Color.Empty;
                }

                // Write the cell text.
                builder.Write(cellData.Text ?? string.Empty);
            }

            // End the current row.
            builder.EndRow();
        }

        // Apply table‑level metadata after at least one row exists.
        if (!string.IsNullOrEmpty(tableData.Title))
            table.Title = tableData.Title;
        if (!string.IsNullOrEmpty(tableData.Description))
            table.Description = tableData.Description;

        // Finish the table.
        builder.EndTable();

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DeserializedTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");
    }
}
