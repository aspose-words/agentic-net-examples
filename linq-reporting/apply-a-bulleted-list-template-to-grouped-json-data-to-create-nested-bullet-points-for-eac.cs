using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample JSON data.
        const string jsonFile = "data.json";
        const string jsonContent = @"{
  ""Categories"": [
    {
      ""Name"": ""Fruits"",
      ""Items"": [ ""Apple"", ""Banana"", ""Cherry"" ]
    },
    {
      ""Name"": ""Vegetables"",
      ""Items"": [ ""Carrot"", ""Lettuce"", ""Tomato"" ]
    }
  ]
}";
        File.WriteAllText(jsonFile, jsonContent);

        // Configure JSON data source to always generate a root object.
        JsonDataLoadOptions loadOptions = new JsonDataLoadOptions
        {
            AlwaysGenerateRootObject = true
        };
        JsonDataSource jsonDataSource = new JsonDataSource(jsonFile, loadOptions);

        // Build the template document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a bulleted list and apply it to the builder.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // Outer foreach – iterate over categories.
        builder.Writeln("<<foreach [cat in data.Categories]>>");
        // Category name – level 0 bullet.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<<[cat.Name]>>");

        // Inner foreach – iterate over items within a category.
        builder.Writeln("<<foreach [itm in cat.Items]>>");
        // Item – level 1 bullet.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("<<[itm]>>");
        // Close inner foreach.
        builder.Writeln("<</foreach>>");

        // Close outer foreach.
        builder.Writeln("<</foreach>>");

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "data");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}
