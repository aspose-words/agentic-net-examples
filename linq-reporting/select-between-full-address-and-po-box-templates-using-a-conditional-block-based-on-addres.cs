using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalAddress
{
    // Root data model for the report.
    public class ReportModel
    {
        public AddressInfo Address { get; set; } = new();
    }

    // Holds address details and the type of address to display.
    public class AddressInfo
    {
        // Expected values: "Full" or "POBox".
        public string Type { get; set; } = string.Empty;

        // Fields for a full street address.
        public string Street { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
        public string State { get; set; } = string.Empty;
        public string Zip { get; set; } = string.Empty;

        // Field for a PO Box address.
        public string POBox { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var modelFull = new ReportModel
            {
                Address = new AddressInfo
                {
                    Type = "Full",
                    Street = "123 Main St.",
                    City = "Springfield",
                    State = "IL",
                    Zip = "62704"
                }
            };

            var modelPoBox = new ReportModel
            {
                Address = new AddressInfo
                {
                    Type = "POBox",
                    POBox = "PO Box 5678"
                }
            };

            // Create a template document with conditional blocks.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Full address block.
            builder.Writeln("<<if [model.Address.Type == \"Full\"]>>");
            builder.Writeln("<<[model.Address.Street]>>");
            builder.Writeln("<<[model.Address.City]>>, <<[model.Address.State]>> <<[model.Address.Zip]>>");
            builder.Writeln("<</if>>");

            // PO Box block.
            builder.Writeln("<<if [model.Address.Type == \"POBox\"]>>");
            builder.Writeln("PO Box: <<[model.Address.POBox]>>");
            builder.Writeln("<</if>>");

            // Build report for a full address.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, modelFull, "model");
            string outputFull = Path.Combine(Environment.CurrentDirectory, "Report_FullAddress.docx");
            template.Save(outputFull);

            // Build report for a PO Box address.
            // Re‑use the same template (it still contains the tags).
            engine.BuildReport(template, modelPoBox, "model");
            string outputPoBox = Path.Combine(Environment.CurrentDirectory, "Report_POBox.docx");
            template.Save(outputPoBox);
        }
    }
}
