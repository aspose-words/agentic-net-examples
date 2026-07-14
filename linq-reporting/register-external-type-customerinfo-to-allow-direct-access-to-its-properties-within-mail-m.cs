using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Sample data model that will be used in the report.
    public class CustomerInfo
    {
        // Initialize properties to avoid nullable warnings.
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string Email { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Customer Report");
            builder.Writeln("Name: <<[customer.Name]>>");
            builder.Writeln("Age: <<[customer.Age]>>");
            builder.Writeln("Email: <<[customer.Email]>>");

            // Save the template to disk before building the report.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            var customer = new CustomerInfo
            {
                Name = "John Doe",
                Age = 30,
                Email = "john.doe@example.com"
            };

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine.
            //    Register the external type so its members can be accessed in the template.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(CustomerInfo));

            // -----------------------------------------------------------------
            // 5. Build the report.
            //    The root object name used in the template is "customer".
            // -----------------------------------------------------------------
            engine.BuildReport(doc, customer, "customer");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
