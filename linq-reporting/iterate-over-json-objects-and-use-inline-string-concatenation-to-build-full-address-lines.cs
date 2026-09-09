using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // ---------- Create sample JSON data ----------
        const string jsonPath = "people.json";
        const string jsonContent = @"[
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
]";
        File.WriteAllText(jsonPath, jsonContent);

        // ---------- Build the template document ----------
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Addresses:");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        // Inline concatenation builds the full address line.
        builder.Writeln("Address: <<[person.Street + \", \" + person.City + \", \" + person.State + \" \" + person.Zip]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // ---------- Load the template ----------
        var reportDoc = new Document(templatePath);

        // ---------- Load JSON data source ----------
        using var jsonStream = File.OpenRead(jsonPath);
        var jsonDataSource = new JsonDataSource(jsonStream);

        // ---------- Generate the report ----------
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // ---------- Save the final report ----------
        reportDoc.Save("Report.docx");
    }
}
