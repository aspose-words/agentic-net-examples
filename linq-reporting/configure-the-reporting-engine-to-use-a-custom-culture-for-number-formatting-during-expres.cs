using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingCustomCulture
{
    // Simple data model used as the root object for the report.
    public class ProductModel
    {
        public string Description { get; set; } = "Sample product";
        public decimal Price { get; set; } = 1234.56m;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required in .NET Core).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare the template document programmatically.
            string templatePath = "Template.docx";
            CreateTemplate(templatePath);

            // Load the template.
            Document doc = new Document(templatePath);

            // Prepare sample data.
            ProductModel model = new ProductModel();

            // Set the thread culture to French for number formatting during expression evaluation.
            Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report.
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            string outputPath = "Report.docx";
            doc.Save(outputPath);
        }

        private static void CreateTemplate(string path)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert LINQ Reporting tags.
            builder.Writeln("Product: <<[model.Description]>>");
            builder.Writeln("Price: <<[model.Price]>>");

            // Save the template.
            doc.Save(path);
        }
    }
}
