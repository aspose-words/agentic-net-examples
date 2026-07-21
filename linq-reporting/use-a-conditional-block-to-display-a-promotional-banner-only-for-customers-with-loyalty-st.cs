using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model for a customer.
    public class Customer
    {
        public string Name { get; set; } = "";
        public bool LoyaltyStatus { get; set; }
    }

    // Wrapper model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Customer Report");
            builder.Writeln(); // Empty line.

            // Begin a foreach loop over the Customers collection.
            builder.Writeln("<<foreach [c in Customers]>>");
            // Output the customer's name.
            builder.Writeln("Name: <<[c.Name]>>");
            // Conditional block: show promotional banner only if LoyaltyStatus is true.
            builder.Writeln("<<if [c.LoyaltyStatus]>>");
            // The banner text is highlighted in red.
            builder.Writeln("<<textColor [\"Red\"]>>*** Special Promotion! ***<</textColor>>");
            builder.Writeln("<</if>>");
            // End of the foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "CustomerReportTemplate.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare sample data.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            var model = new ReportModel
            {
                Customers = new List<Customer>
                {
                    new Customer { Name = "Alice",   LoyaltyStatus = true  },
                    new Customer { Name = "Bob",     LoyaltyStatus = false },
                    new Customer { Name = "Charlie", LoyaltyStatus = true  }
                }
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "model".
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "CustomerReportOutput.docx";
            reportDoc.Save(outputPath);
        }
    }
}
