using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Text.Json;

// Define classes that match the JSON structure.
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

// Main routine that builds the Word document from JSON.
public class JsonTableToWord
{
    public static void BuildDocument(string jsonPath, string outputDocPath)
    {
        // Load JSON from file.
        string json = File.ReadAllText(jsonPath);
        // Use System.Text.Json for deserialization (no external NuGet needed).
        var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        TableData tableData = JsonSerializer.Deserialize<TableData>(json, options);

        // Create a blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Apply optional table‑level formatting.
        if (!string.IsNullOrEmpty(tableData.TableStyle))
        {
            if (Enum.TryParse(tableData.TableStyle, out StyleIdentifier styleId))
                table.StyleIdentifier = styleId;
            else
                table.StyleName = tableData.TableStyle; // fallback to name.
        }

        if (tableData.PreferredWidth.HasValue)
            table.PreferredWidth = PreferredWidth.FromPoints(tableData.PreferredWidth.Value);

        if (tableData.AllowAutoFit.HasValue)
            table.AllowAutoFit = tableData.AllowAutoFit.Value;

        if (!string.IsNullOrEmpty(tableData.Title))
            table.Title = tableData.Title;

        if (!string.IsNullOrEmpty(tableData.Description))
            table.Description = tableData.Description;

        // Populate rows and cells.
        foreach (RowData rowData in tableData.Rows)
        {
            // For each cell in the current row.
            foreach (string cellText in rowData.Cells)
            {
                builder.InsertCell();
                builder.Write(cellText);
            }
            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document.
        doc.Save(outputDocPath);
    }

    // Example usage.
    public static void Main()
    {
        // Path to the JSON file containing table data.
        string jsonFile = @"C:\Data\sampleTable.json";

        // Desired output Word document path.
        string outputFile = @"C:\Output\GeneratedTable.docx";

        BuildDocument(jsonFile, outputFile);

        Console.WriteLine("Document created successfully.");
    }
}
