using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample JSON data file.
        string jsonPath = "data.json";
        string sampleJson = @"{
  ""Products"": [
    { ""Name"": ""Apple"",  ""Price"": 1.23 },
    { ""Name"": ""Banana"", ""Price"": 0.99 },
    { ""Name"": ""Cherry"", ""Price"": 2.50 }
  ]
}";
        File.WriteAllText(jsonPath, sampleJson);

        // Deserialize JSON into the data model.
        string jsonContent = File.ReadAllText(jsonPath);
        ReportModel model = JsonConvert.DeserializeObject<ReportModel>(jsonContent) ?? new ReportModel();

        // Create the LINQ Reporting template document.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Report");
        builder.Writeln();

        // Begin foreach loop over Products collection.
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        // Use a pre‑formatted property for currency output.
        builder.Writeln("Price: <<[p.FormattedPrice]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Build the report using the data model.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

// Product item model.
public class Product
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }

    // Returns the price formatted as currency using the current culture.
    public string FormattedPrice => Price.ToString("C", CultureInfo.CurrentCulture);
}
