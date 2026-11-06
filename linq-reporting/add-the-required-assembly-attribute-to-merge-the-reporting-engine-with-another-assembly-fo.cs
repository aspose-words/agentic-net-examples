using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingAssemblyMergeExample
{
    // Simple data model used by the template.
    public class ReportModel
    {
        public string Name { get; set; } = "World";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string outputPath = "Report.docx";

            // -----------------------------------------------------------------
            // Create a Word document template with a LINQ Reporting tag.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Hello <<[model.Name]>>!");
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);
            var model = new ReportModel(); // model.Name is already initialized.

            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save(outputPath);
        }
    }
}
