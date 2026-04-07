using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample JSON data file.
        // -----------------------------------------------------------------
        string jsonPath = "data.json";
        string jsonContent = @"{
  ""Categories"": [
    {
      ""Name"": ""Fruits"",
      ""Items"": [ ""Apple"", ""Banana"", ""Cherry"" ]
    },
    {
      ""Name"": ""Vegetables"",
      ""Items"": [ ""Carrot"", ""Lettuce"" ]
    },
    {
      ""Name"": ""Beverages"",
      ""Items"": [ ""Water"", ""Coffee"", ""Tea"" ]
    }
  ]
}";
        File.WriteAllText(jsonPath, jsonContent);

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a bulleted list and apply it to the builder.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // Template starts – iterate over categories.
        builder.Writeln("<<foreach [cat in Categories]>>");

        // Category name – top‑level bullet.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<<[cat.Name]>>");

        // Items inside the current category – second‑level bullets.
        builder.Writeln("<<foreach [itm in cat.Items]>>");
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("<<[itm]>>");
        builder.Writeln("<</foreach>>"); // end inner foreach

        builder.Writeln("<</foreach>>"); // end outer foreach

        // -----------------------------------------------------------------
        // 3. Load JSON data as a LINQ Reporting data source.
        // -----------------------------------------------------------------
        JsonDataLoadOptions loadOptions = new JsonDataLoadOptions
        {
            // Ensure the root object is generated so that "Categories" can be accessed.
            AlwaysGenerateRootObject = true
        };
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath, loadOptions);

        // -----------------------------------------------------------------
        // 4. Build the report.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource); // root object is the JSON root

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
