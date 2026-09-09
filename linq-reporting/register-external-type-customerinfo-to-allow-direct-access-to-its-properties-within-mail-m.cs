using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Sample data model.
    public class CustomerInfo
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
        public string Email { get; set; } = "john.doe@example.com";
    }

    // Wrapper for the root object used in the template.
    public class ReportModel
    {
        public CustomerInfo Customer { get; set; } = new CustomerInfo();
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document and a builder to insert LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert template tags that reference the CustomerInfo properties.
            builder.Writeln("Customer Report");
            builder.Writeln("----------------");
            builder.Writeln("Name : <<[model.Customer.Name]>>");
            builder.Writeln("Age  : <<[model.Customer.Age]>>");
            builder.Writeln("Email: <<[model.Customer.Email]>>");

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Register the external type so its members can be accessed in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(CustomerInfo));

            // Build the report using the template, data source, and root name.
            engine.BuildReport(doc, model, "model");

            // Save the generated document.
            doc.Save("CustomerReport.docx");
        }
    }
}
