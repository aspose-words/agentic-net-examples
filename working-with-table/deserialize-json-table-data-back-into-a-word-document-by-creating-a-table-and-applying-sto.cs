using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    // Classes that represent the JSON structure for a table.
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
        public string BackgroundColor { get; set; } // Color name or hex string.
    }

    public static void Main()
    {
        // Sample JSON that stores table content and simple cell formatting.
        string json = @"
        {
            ""Rows"": [
                {
                    ""Cells"": [
                        { ""Text"": ""Header 1"", ""BackgroundColor"": ""LightGray"" },
                        { ""Text"": ""Header 2"", ""BackgroundColor"": ""LightGray"" }
                    ]
                },
                {
                    ""Cells"": [
                        { ""Text"": ""Row1 Col1"", ""BackgroundColor"": ""White"" },
                        { ""Text"": ""Row1 Col2"", ""BackgroundColor"": ""White"" }
                    ]
                },
                {
                    ""Cells"": [
                        { ""Text"": ""Row2 Col1"", ""BackgroundColor"": ""White"" },
                        { ""Text"": ""Row2 Col2"", ""BackgroundColor"": ""White"" }
                    ]
                }
            ]
        }";

        // Deserialize JSON into C# objects.
        TableData tableData = JsonConvert.DeserializeObject<TableData>(json);
        if (tableData == null || tableData.Rows == null)
            throw new InvalidOperationException("Failed to deserialize table data.");

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Optional: add a title before the table.
        builder.Writeln("Table generated from JSON data:");
        builder.Writeln();

        // Build the table using DocumentBuilder.
        builder.StartTable();

        foreach (RowData row in tableData.Rows)
        {
            if (row.Cells == null)
                continue;

            foreach (CellData cell in row.Cells)
            {
                // Insert a new cell.
                builder.InsertCell();

                // Write the cell text.
                builder.Write(cell.Text ?? string.Empty);

                // Apply background shading if a color is specified.
                if (!string.IsNullOrEmpty(cell.BackgroundColor))
                {
                    Color bgColor = Color.FromName(cell.BackgroundColor);
                    // If the name is not known, try parsing as a hex value.
                    if (!bgColor.IsKnownColor && !bgColor.IsNamedColor)
                    {
                        try
                        {
                            bgColor = ColorTranslator.FromHtml(cell.BackgroundColor);
                        }
                        catch
                        {
                            // Fallback to white if parsing fails.
                            bgColor = Color.White;
                        }
                    }

                    builder.CurrentParagraph.ParagraphFormat.Shading.BackgroundPatternColor = bgColor;
                }
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
