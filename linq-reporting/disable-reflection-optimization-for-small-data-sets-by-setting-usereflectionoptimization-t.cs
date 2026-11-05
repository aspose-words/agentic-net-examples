using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingReflectionOptimization
{
    // Sample data model
    public class Customer
    {
        public string Name { get; set; } = "John Doe";
    }

    public class ReportModel
    {
        public Customer Customer { get; set; } = new Customer();
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure code page support (required by Aspose.Words for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // Create a simple template document with a LINQ Reporting tag
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Customer Name: <<[model.Customer.Name]>>");
            templateDoc.Save(templatePath);

            // Load the template document
            Document doc = new Document(templatePath);

            // Prepare the data model
            ReportModel model = new ReportModel();

            // Disable reflection optimization for small data sets
            ReportingEngine.UseReflectionOptimization = false;

            // Build the report
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            doc.Save(outputPath);
        }
    }
}
