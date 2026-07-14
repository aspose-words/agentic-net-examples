using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
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
            // Step 1: Create a DOCX template with a LINQ Reporting tag.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("Hello <<[model.Name]>>!");
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Step 2: Load the template from the file system.
            Document doc = new Document(templatePath);

            // Step 3: Prepare the data source.
            ReportModel model = new ReportModel { Name = "Aspose.Words" };

            // Step 4: Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Step 5: Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
