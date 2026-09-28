using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingReflectionOptimization
{
    // Simple data model used by the template.
    public class Model
    {
        public string Name { get; set; } = "World";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare folders.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // Create a template document with a LINQ Reporting tag.
            string templatePath = Path.Combine(outputDir, "template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Hello <<[model.Name]>>!");
            templateDoc.Save(templatePath);

            // Load the template for report generation.
            Document loadedTemplate = new Document(templatePath);

            // Prepare the data source.
            Model model = new Model();

            // Disable reflection optimization only for this report generation.
            bool previousOptimizationSetting = ReportingEngine.UseReflectionOptimization;
            ReportingEngine.UseReflectionOptimization = false;
            try
            {
                ReportingEngine engine = new ReportingEngine();
                engine.BuildReport(loadedTemplate, model, "model");
            }
            finally
            {
                // Restore the original setting.
                ReportingEngine.UseReflectionOptimization = previousOptimizationSetting;
            }

            // Save the generated report.
            string resultPath = Path.Combine(outputDir, "result.docx");
            loadedTemplate.Save(resultPath);
        }
    }
}
