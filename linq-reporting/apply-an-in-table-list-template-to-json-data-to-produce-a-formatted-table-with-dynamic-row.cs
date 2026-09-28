using System;
using System.Collections.Generic;
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
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Prepare sample JSON data.
        // -----------------------------------------------------------------
        string jsonPath = "data.json";
        var sampleJson = @"{
  ""Items"": [
    { ""Index"": 1, ""Name"": ""Item A"", ""Value"": ""100"" },
    { ""Index"": 2, ""Name"": ""Item B"", ""Value"": ""200"" },
    { ""Index"": 3, ""Name"": ""Item C"", ""Value"": ""300"" }
  ]
}";
        File.WriteAllText(jsonPath, sampleJson, Encoding.UTF8);

        // Deserialize JSON into a strongly‑typed model.
        var model = JsonConvert.DeserializeObject<ReportModel>(File.ReadAllText(jsonPath, Encoding.UTF8))!;

        // -----------------------------------------------------------------
        // 2. Build the template document.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Header table (static – appears once).
        // -----------------------------------------------------------------
        Table headerTable = builder.StartTable();

        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Value");
        builder.EndRow();

        builder.EndTable();

        // -----------------------------------------------------------------
        // Row template – placed inside a foreach block.
        // -----------------------------------------------------------------
        builder.Writeln("<<foreach [item in Items]>>");

        Table rowTable = builder.StartTable();

        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Value]>>");
        builder.EndRow();

        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // -----------------------------------------------------------------
        // 3. Build the report.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated document.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine("Output", "Report.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        doc.Save(outputPath);

        Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
    }
}

// ---------------------------------------------------------------------
// Data model aligned with the JSON structure.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
    public string Value { get; set; } = "";
}
