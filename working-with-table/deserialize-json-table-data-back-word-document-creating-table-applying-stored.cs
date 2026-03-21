using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Tables;

public class TableData
{
    public List<RowData> Rows { get; set; }
    public string TableStyle { get; set; }          // e.g., "MediumShading1Accent1"
    public string Title { get; set; }               // optional
    public string Description { get; set; }         // optional
    public double? PreferredWidth { get; set; }     // optional, in points
    public bool? AllowAutoFit { get; set; }         // optional
}

public class RowData
{
    public List<string> Cells { get; set; }
}

public class JsonTableToWord
{
    public static void BuildDocument(string jsonPath, string outputDocPath)
    {
        string json = File.ReadAllText(jsonPath);
        var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        TableData tableData = JsonSerializer.Deserialize<TableData>(json, options);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();

        // Build rows and cells
        if (tableData?.Rows != null)
        {
            foreach (RowData rowData in tableData.Rows)
            {
                foreach (string cellText in rowData.Cells)
                {
                    builder.InsertCell();
                    builder.Write(cellText);
                }
                builder.EndRow();
            }
        }

        // Apply formatting after the table has at least one row
        if (!string.IsNullOrEmpty(tableData?.TableStyle))
        {
            if (Enum.TryParse(tableData.TableStyle, out StyleIdentifier styleId))
                table.StyleIdentifier = styleId;
            else
                table.StyleName = tableData.TableStyle;
        }

        if (tableData?.PreferredWidth.HasValue == true)
            table.PreferredWidth = PreferredWidth.FromPoints(tableData.PreferredWidth.Value);

        if (tableData?.AllowAutoFit.HasValue == true)
            table.AllowAutoFit = tableData.AllowAutoFit.Value;

        if (!string.IsNullOrEmpty(tableData?.Title))
            table.Title = tableData.Title;

        if (!string.IsNullOrEmpty(tableData?.Description))
            table.Description = tableData.Description;

        builder.EndTable();
        doc.Save(outputDocPath);
    }

    public static void Main()
    {
        string baseDir = AppContext.BaseDirectory;
        string jsonFile = Path.Combine(baseDir, "sampleTable.json");
        string outputFile = Path.Combine(baseDir, "GeneratedTable.docx");

        if (!File.Exists(jsonFile))
        {
            var sampleData = new TableData
            {
                TableStyle = "MediumShading1Accent1",
                Title = "Sample Table",
                Description = "A table generated from JSON data.",
                PreferredWidth = 500,
                AllowAutoFit = true,
                Rows = new List<RowData>
                {
                    new RowData { Cells = new List<string> { "Header 1", "Header 2", "Header 3" } },
                    new RowData { Cells = new List<string> { "Row 1, Cell 1", "Row 1, Cell 2", "Row 1, Cell 3" } },
                    new RowData { Cells = new List<string> { "Row 2, Cell 1", "Row 2, Cell 2", "Row 2, Cell 3" } }
                }
            };

            string json = JsonSerializer.Serialize(sampleData, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(jsonFile, json);
        }

        BuildDocument(jsonFile, outputFile);
        Console.WriteLine($"Document created successfully at: {outputFile}");
    }
}
