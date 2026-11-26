using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json.Linq;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words in some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "data.json");
        string jsonContent = @"[
  { ""Name"": ""Alice"", ""Age"": 30, ""City"": ""New York"" },
  { ""Name"": ""Bob"",   ""Age"": 25, ""City"": ""Los Angeles"" },
  { ""Name"": ""Charlie"", ""Age"": 28, ""City"": ""Chicago"" }
]";
        File.WriteAllText(jsonPath, jsonContent);

        // Parse JSON to obtain column headers (property names)
        JArray jsonArray = JArray.Parse(File.ReadAllText(jsonPath));
        List<string> headers = new();
        if (jsonArray.Count > 0 && jsonArray[0] is JObject firstObj)
        {
            foreach (var prop in firstObj.Properties())
                headers.Add(prop.Name);
        }

        // Create the template document
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin foreach loop over JSON data rows
        builder.Writeln("<<foreach [row in data]>>");

        // Start a table for each iteration (header will be repeated for simplicity)
        Table table = builder.StartTable();

        // Header row
        foreach (string header in headers)
        {
            builder.InsertCell();
            builder.Writeln(header);
        }
        builder.EndRow();

        // Data row
        foreach (string header in headers)
        {
            builder.InsertCell();
            // Tag that will be replaced with the corresponding field value
            builder.Writeln($"<<[row.{header}]>>");
        }
        builder.EndRow();

        // End the table and the foreach block
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Build the report using the JSON data source
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(template, jsonDataSource, "data");

        // Save the generated report
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        template.Save(outputPath);
    }
}
