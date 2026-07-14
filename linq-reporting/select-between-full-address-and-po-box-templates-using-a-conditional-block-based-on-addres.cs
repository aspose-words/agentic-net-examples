using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model for an address.
    public class Address
    {
        // Determines which template block to use.
        public bool IsPoBox { get; set; } = false;

        // PO Box address fields.
        public string PoBox { get; set; } = string.Empty;

        // Full street address fields.
        public string Street { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
        public string State { get; set; } = string.Empty;
        public string Zip { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare two sample data objects.
            var poBoxAddress = new Address
            {
                IsPoBox = true,
                PoBox = "PO Box 1234"
            };

            var fullAddress = new Address
            {
                IsPoBox = false,
                Street = "123 Main St.",
                City = "Springfield",
                State = "IL",
                Zip = "62704"
            };

            // Create a template document programmatically.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Recipient Address:");
            // PO Box block – shown only when IsPoBox is true.
            builder.Writeln("<<if [model.IsPoBox]>>");
            builder.Writeln("<<[model.PoBox]>>");
            builder.Writeln("<</if>>");
            // Full address block – shown only when IsPoBox is false.
            builder.Writeln("<<if [!model.IsPoBox]>>");
            builder.Writeln("<<[model.Street]>>");
            builder.Writeln("<<[model.City]>>, <<[model.State]>> <<[model.Zip]>>");
            builder.Writeln("<</if>>");

            // Build the report for the PO Box address.
            var engine = new ReportingEngine();
            engine.BuildReport(template, poBoxAddress, "model");
            template.Save("Report_POBox.docx");

            // Re‑use the same template for the full address.
            // (Reloading the template ensures a clean state.)
            var templateForFull = new Document("Report_POBox.docx");
            var engineFull = new ReportingEngine();
            engineFull.BuildReport(templateForFull, fullAddress, "model");
            templateForFull.Save("Report_Full.docx");
        }
    }
}
