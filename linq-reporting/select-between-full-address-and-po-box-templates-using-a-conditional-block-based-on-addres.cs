using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalAddress
{
    // Data model used by the LINQ Reporting engine.
    public class AddressModel
    {
        // Indicates whether the address is a PO Box.
        public bool IsPoBox { get; set; } = false;

        // Full street address (used when IsPoBox is false).
        public string FullAddress { get; set; } = "123 Main St., Springfield, USA";

        // PO Box address (used when IsPoBox is true).
        public string PoBox { get; set; } = "PO Box 456, Springfield, USA";
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Conditional block: if the address is a PO Box, use the PO Box template.
            builder.Writeln("<<if [model.IsPoBox]>>");
            builder.Writeln("<<[model.PoBox]>>");
            builder.Writeln("<</if>>");

            // Conditional block: if the address is NOT a PO Box, use the full address template.
            builder.Writeln("<<if [!model.IsPoBox]>>");
            builder.Writeln("<<[model.FullAddress]>>");
            builder.Writeln("<</if>>");

            // Save the template to disk.
            const string templatePath = "AddressTemplate.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // Sample data: first example uses a PO Box address.
            AddressModel poBoxModel = new AddressModel
            {
                IsPoBox = true,
                FullAddress = "Should not appear",
                PoBox = "PO Box 789, Metropolis, USA"
            };

            // Sample data: second example uses a full street address.
            AddressModel fullAddressModel = new AddressModel
            {
                IsPoBox = false,
                FullAddress = "742 Evergreen Terrace, Springfield, USA",
                PoBox = "Should not appear"
            };

            // Build report for PO Box scenario.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, poBoxModel, "model");
            doc.Save("Report_POBox.docx");

            // Reload the template for the second scenario.
            doc = new Document(templatePath);
            engine.BuildReport(doc, fullAddressModel, "model");
            doc.Save("Report_FullAddress.docx");
        }
    }
}
