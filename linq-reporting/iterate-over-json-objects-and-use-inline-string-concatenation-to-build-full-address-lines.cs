using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string json = @"{
  ""Persons"": [
    {
      ""Name"": ""John Doe"",
      ""Street"": ""123 Main St"",
      ""City"": ""Springfield"",
      ""State"": ""IL"",
      ""Zip"": ""62704""
    },
    {
      ""Name"": ""Jane Smith"",
      ""Street"": ""456 Oak Ave"",
      ""City"": ""Metropolis"",
      ""State"": ""NY"",
      ""Zip"": ""10001""
    }
  ]
}";
        // Write JSON to a local file.
        string dataPath = Path.Combine(Environment.CurrentDirectory, "data.json");
        File.WriteAllText(dataPath, json, Encoding.UTF8);

        // Deserialize JSON into the model.
        ReportModel model = JsonConvert.DeserializeObject<ReportModel>(json) ?? new();

        // Create the template document programmatically.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Address: <<[p.Street + \", \" + p.City + \", \" + p.State + \" \" + p.Zip]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new(templatePath);

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "report.docx");
        reportDoc.Save(outputPath);
    }
}

// Root wrapper class for the JSON data.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Person data class.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public string Street { get; set; } = string.Empty;
    public string City { get; set; } = string.Empty;
    public string State { get; set; } = string.Empty;
    public string Zip { get; set; } = string.Empty;
}
