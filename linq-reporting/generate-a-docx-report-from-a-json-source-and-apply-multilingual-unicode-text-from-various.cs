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
        // Register code page provider for Aspose.Words if needed
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data with multilingual fields
        string json = @"
{
  ""Items"": [
    {
      ""Id"": 1,
      ""NameEn"": ""Apple"",
      ""NameRu"": ""Яблоко"",
      ""NameZh"": ""苹果"",
      ""NameAr"": ""تفاح""
    },
    {
      ""Id"": 2,
      ""NameEn"": ""Banana"",
      ""NameRu"": ""Банан"",
      ""NameZh"": ""香蕉"",
      ""NameAr"": ""موز""
    },
    {
      ""Id"": 3,
      ""NameEn"": ""Cherry"",
      ""NameRu"": ""Вишня"",
      ""NameZh"": ""樱桃"",
      ""NameAr"": ""كرز""
    }
  ]
}";
        // Write JSON to a file (optional, just to demonstrate file usage)
        File.WriteAllText("data.json", json, Encoding.UTF8);

        // Deserialize JSON into model
        ReportData model = JsonConvert.DeserializeObject<ReportData>(json) ?? new ReportData();

        // Create a template document programmatically
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Multilingual Items Report");
        builder.Writeln("");

        // Begin foreach over Items
        builder.Writeln("<<foreach [item in Items]>>");

        // Write item ID
        builder.Writeln("Item ID: <<[item.Id]>>");

        // Create a table for the multilingual names
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell(); builder.Writeln("English");
        builder.InsertCell(); builder.Writeln("Russian");
        builder.InsertCell(); builder.Writeln("Chinese");
        builder.InsertCell(); builder.Writeln("Arabic");
        builder.EndRow();

        // Data row with tags
        builder.InsertCell(); builder.Writeln("<<[item.NameEn]>>");
        builder.InsertCell(); builder.Writeln("<<[item.NameRu]>>");
        builder.InsertCell(); builder.Writeln("<<[item.NameZh]>>");
        builder.InsertCell(); builder.Writeln("<<[item.NameAr]>>");
        builder.EndRow();

        builder.EndTable();

        // End foreach
        builder.Writeln("<</foreach>>");

        // Save the template
        template.Save("template.docx");

        // Load the template for reporting
        Document reportDoc = new Document("template.docx");

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "data");

        // Save the final report
        reportDoc.Save("report.docx");
    }

    // Data model classes
    public class ReportData
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Id { get; set; }
        public string NameEn { get; set; } = "";
        public string NameRu { get; set; } = "";
        public string NameZh { get; set; } = "";
        public string NameAr { get; set; } = "";
    }
}
