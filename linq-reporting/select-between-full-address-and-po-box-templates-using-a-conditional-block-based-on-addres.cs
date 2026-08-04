using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model for the report.
    public class ReportModel
    {
        public Address Address { get; set; } = new();
    }

    // Address information with a flag indicating PO Box usage.
    public class Address
    {
        public bool IsPoBox { get; set; }
        public string FullAddress { get; set; } = "";
        public string PoBox { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Conditional block: if the address is a PO Box, show PO Box field; otherwise show full address.
            builder.Writeln("<<if [model.Address.IsPoBox]>>");
            builder.Writeln("PO Box: <<[model.Address.PoBox]>>");
            builder.Writeln("<</if>>");
            builder.Writeln("<<if [model.Address.IsPoBox == false]>>");
            builder.Writeln("Address: <<[model.Address.FullAddress]>>");
            builder.Writeln("<</if>>");

            // First example: a PO Box address.
            ReportModel poBoxModel = new ReportModel
            {
                Address = new Address
                {
                    IsPoBox = true,
                    PoBox = "PO Box 1234",
                    FullAddress = "123 Main St, Springfield"
                }
            };

            // Build the report for the PO Box scenario.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, poBoxModel, "model");
            template.Save("Report_POBox.docx");

            // Second example: a regular full address.
            Document template2 = new Document();
            DocumentBuilder builder2 = new DocumentBuilder(template2);
            builder2.Writeln("<<if [model.Address.IsPoBox]>>");
            builder2.Writeln("PO Box: <<[model.Address.PoBox]>>");
            builder2.Writeln("<</if>>");
            builder2.Writeln("<<if [model.Address.IsPoBox == false]>>");
            builder2.Writeln("Address: <<[model.Address.FullAddress]>>");
            builder2.Writeln("<</if>>");

            ReportModel fullAddressModel = new ReportModel
            {
                Address = new Address
                {
                    IsPoBox = false,
                    PoBox = "PO Box 9999",
                    FullAddress = "456 Oak Avenue, Metropolis"
                }
            };

            engine.BuildReport(template2, fullAddressModel, "model");
            template2.Save("Report_FullAddress.docx");
        }
    }
}
