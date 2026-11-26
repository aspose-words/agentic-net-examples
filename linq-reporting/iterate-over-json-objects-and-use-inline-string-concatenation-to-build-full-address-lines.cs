using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

namespace LinqReportingJsonAddress
{
    // Root model for the report.
    public class ReportModel
    {
        public List<AddressItem> Addresses { get; set; } = new();
    }

    // Individual address item.
    public class AddressItem
    {
        public string Street { get; set; } = "";
        public string City { get; set; } = "";
        public string State { get; set; } = "";
        public string Zip { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create sample JSON data file.
            const string jsonFile = "addresses.json";
            var sampleJson = @"[
  { ""Street"": ""123 Main St"", ""City"": ""Springfield"", ""State"": ""IL"", ""Zip"": ""62704"" },
  { ""Street"": ""456 Oak Ave"", ""City"": ""Metropolis"", ""State"": ""NY"", ""Zip"": ""10001"" },
  { ""Street"": ""789 Pine Rd"", ""City"": ""Gotham"", ""State"": ""NJ"", ""Zip"": ""07097"" }
]";
            File.WriteAllText(jsonFile, sampleJson);

            // Load JSON into C# objects.
            var addressList = JsonConvert.DeserializeObject<List<AddressItem>>(File.ReadAllText(jsonFile)) ?? new();
            var model = new ReportModel { Addresses = addressList };

            // Create the template document programmatically.
            const string templateFile = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Address Report");
            builder.Writeln("<<foreach [addr in model.Addresses]>>");
            // Inline string concatenation to build full address line.
            builder.Writeln("<<[addr.Street + \", \" + addr.City + \", \" + addr.State + \" \" + addr.Zip]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templateFile);

            // Load the template for report generation.
            var reportDoc = new Document(templateFile);

            // Build the report using LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the final report.
            const string outputFile = "Report.docx";
            reportDoc.Save(outputFile);
        }
    }
}
