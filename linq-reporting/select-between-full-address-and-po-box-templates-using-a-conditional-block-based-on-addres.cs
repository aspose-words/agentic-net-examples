using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalAddress
{
    // Data model classes
    public class ReportModel
    {
        public Address Address { get; set; } = new();
    }

    public class Address
    {
        public bool IsPoBox { get; set; }
        public string PoBoxNumber { get; set; } = string.Empty;
        public string Street { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
        public string State { get; set; } = string.Empty;
        public string Zip { get; set; } = string.Empty;
    }

    class Program
    {
        static void Main()
        {
            // Ensure output directory exists
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create the template document programmatically
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a title
            builder.Writeln("Address Report");
            builder.Writeln();

            // Conditional block for PO Box
            builder.Writeln("<<if [model.Address.IsPoBox]>>");
            builder.Writeln("PO Box: <<[model.Address.PoBoxNumber]>>");
            builder.Writeln("<</if>>");

            // Conditional block for full street address
            builder.Writeln("<<if [!model.Address.IsPoBox]>>");
            builder.Writeln("Street: <<[model.Address.Street]>>");
            builder.Writeln("City: <<[model.Address.City]>>");
            builder.Writeln("State: <<[model.Address.State]>>");
            builder.Writeln("ZIP: <<[model.Address.Zip]>>");
            builder.Writeln("<</if>>");

            // Save the template to disk
            string templatePath = Path.Combine(outputDir, "AddressTemplate.docx");
            templateDoc.Save(templatePath);

            // 2. Prepare sample data: one with PO Box, one with full address
            var poBoxModel = new ReportModel
            {
                Address = new Address
                {
                    IsPoBox = true,
                    PoBoxNumber = "PO Box 1234"
                }
            };

            var fullAddressModel = new ReportModel
            {
                Address = new Address
                {
                    IsPoBox = false,
                    Street = "123 Main St",
                    City = "Springfield",
                    State = "IL",
                    Zip = "62704"
                }
            };

            // 3. Generate reports using ReportingEngine
            ReportingEngine engine = new ReportingEngine();

            // Load the template document
            Document docForPoBox = new Document(templatePath);
            bool successPoBox = engine.BuildReport(docForPoBox, poBoxModel, "model");
            string poBoxOutput = Path.Combine(outputDir, "PoBoxReport.docx");
            docForPoBox.Save(poBoxOutput);

            Document docForFull = new Document(templatePath);
            bool successFull = engine.BuildReport(docForFull, fullAddressModel, "model");
            string fullOutput = Path.Combine(outputDir, "FullAddressReport.docx");
            docForFull.Save(fullOutput);

            // Optionally, indicate success (no interactive output required)
            // The example finishes here.
        }
    }
}
