using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingReflectionOptimization
{
    // Simple data model used by the template.
    public class Model
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a template document with a LINQ Reporting tag.
            var template = new Document();
            var builder = new DocumentBuilder(template);
            builder.Writeln("Customer name: <<[model.Name]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template back for reporting.
            var doc = new Document(templatePath);

            // Disable reflection optimization for small data sets.
            ReportingEngine.UseReflectionOptimization = false;

            // Prepare the data source.
            var model = new Model();

            // Build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
