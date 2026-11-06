using System;
using System.Data;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // ---------- Step 1: Create sample JSON data ----------
        string json = @"[
            { ""Name"": ""Alice"", ""Age"": 30, ""City"": ""New York"" },
            { ""Name"": ""Bob"",   ""Age"": 25, ""City"": ""Los Angeles"" },
            { ""Name"": ""Carol"", ""Age"": 28, ""City"": ""Chicago"" }
        ]";

        const string jsonPath = "data.json";
        File.WriteAllText(jsonPath, json);

        // ---------- Step 2: Load JSON into a DataTable ----------
        DataTable table = JsonConvert.DeserializeObject<DataTable>(File.ReadAllText(jsonPath))!;

        // Wrap the DataTable rows for the reporting engine.
        var model = new ReportModel
        {
            TableRows = table.Rows
        };

        // ---------- Step 3: Create the LINQ Reporting template ----------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin foreach loop over rows.
        builder.Writeln("<<foreach [row in TableRows]>>");

        // Start the table.
        Table reportTable = builder.StartTable();

        // Header row: column names derived from the DataTable schema.
        foreach (DataColumn column in table.Columns)
        {
            builder.InsertCell();
            builder.Writeln(column.ColumnName);
        }
        builder.EndRow();

        // Data row: each cell contains a field tag referencing the column name via indexer.
        foreach (DataColumn column in table.Columns)
        {
            builder.InsertCell();
            // Use the DataRow indexer syntax: row["ColumnName"]
            builder.Writeln($"<<[row[\"{column.ColumnName}\"]]>>");
        }
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection).
        const string templatePath = "template.docx";
        doc.Save(templatePath);

        // ---------- Step 4: Build the report ----------
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // ---------- Step 5: Save the generated report ----------
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}

// Wrapper class exposing the rows collection to the template.
public class ReportModel
{
    public DataRowCollection TableRows { get; set; } = null!;
}
