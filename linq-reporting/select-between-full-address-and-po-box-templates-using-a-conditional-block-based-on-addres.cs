using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model representing an address.
    public class Address
    {
        // Determines which template block to use.
        public bool IsPoBox { get; set; }

        // Full street address (used when IsPoBox is false).
        public string FullAddress { get; set; } = string.Empty;

        // PO Box address (used when IsPoBox is true).
        public string PoBox { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for older encodings (required by Aspose.Words in some environments).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Path for the temporary template file.
            const string templatePath = "AddressTemplate.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document with conditional LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Conditional block for PO Box addresses.
            builder.Writeln("<<if [address.IsPoBox]>>");
            builder.Writeln("PO Box: <<[address.PoBox]>>");
            builder.Writeln("<</if>>");

            // Conditional block for full street addresses.
            builder.Writeln("<<if [address.IsPoBox == false]>>");
            builder.Writeln("Address: <<[address.FullAddress]>>");
            builder.Writeln("<</if>>");

            // Save the template so it can be loaded later.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data – one full address and one PO Box address.
            // -----------------------------------------------------------------
            var addresses = new List<Address>
            {
                new Address
                {
                    IsPoBox = false,
                    FullAddress = "123 Main St, Springfield, USA"
                },
                new Address
                {
                    IsPoBox = true,
                    PoBox = "PO Box 456, Springfield, USA"
                }
            };

            // -----------------------------------------------------------------
            // 3. Generate a report for each address using the ReportingEngine.
            // -----------------------------------------------------------------
            int index = 1;
            foreach (var address in addresses)
            {
                // Load the template.
                Document reportDoc = new Document(templatePath);

                // Build the report. The root object name in the template is "address".
                ReportingEngine engine = new ReportingEngine();
                engine.BuildReport(reportDoc, address, "address");

                // Save the generated report.
                string outputFileName = $"AddressReport_{index}.docx";
                reportDoc.Save(outputFileName);
                Console.WriteLine($"Report saved: {Path.GetFullPath(outputFileName)}");
                index++;
            }

            // Clean up the temporary template file (optional).
            if (File.Exists(templatePath))
                File.Delete(templatePath);
        }
    }
}
