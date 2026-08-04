using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Ensure the output folder exists.
        Directory.CreateDirectory("Output");

        // 1. Create a simple JSON file with grouped data.
        const string jsonPath = "Output/Data.json";
        File.WriteAllText(jsonPath,
            @"{
  ""Categories"": [
    {
      ""Name"": ""Fruits"",
      ""Items"": [ ""Apple"", ""Banana"", ""Cherry"" ]
    },
    {
      ""Name"": ""Vegetables"",
      ""Items"": [ ""Carrot"", ""Lettuce"" ]
    }
  ]
}");

        // 2. Deserialize JSON into a strongly‑typed model.
        RootModel model = JsonConvert.DeserializeObject<RootModel>(File.ReadAllText(jsonPath))!;

        // 3. Build the LINQ Reporting template programmatically.
        const string templatePath = "Output/Template.docx";
        CreateTemplate(templatePath);

        // 4. Load the template and run the reporting engine.
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated report.
        const string reportPath = "Output/Report.docx";
        doc.Save(reportPath);
    }

    // Creates a Word document that contains LINQ Reporting tags and a bulleted list style.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a bulleted list that will be used for both levels.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Begin outer loop over categories.
        builder.Writeln("<<foreach [cat in Categories]>>");

        // Category name – level 0 bullet.
        builder.ListFormat.List = bulletList;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<<[cat.Name]>>");

        // Begin inner loop over items.
        builder.Writeln("<<foreach [item in cat.Items]>>");

        // Item – level 1 bullet.
        builder.ListFormat.List = bulletList;
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("<<[item]>>");

        // End inner loop.
        builder.Writeln("<</foreach>>");

        // End outer loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }

    // Root object that matches the JSON structure.
    public class RootModel
    {
        public List<Category> Categories { get; set; } = new();
    }

    // Category with a name and a collection of items.
    public class Category
    {
        public string Name { get; set; } = string.Empty;
        public List<string> Items { get; set; } = new();
    }
}
