using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Prepare file paths.
            string workDir = Directory.GetCurrentDirectory();
            string templatePath = Path.Combine(workDir, "template.docx");
            string resultPath = Path.Combine(workDir, "result.docx");

            // -----------------------------------------------------------------
            // 1. Create a simple template document with a LINQ Reporting tag.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Hello <<[model.Name]>>!");
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for reporting.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data model.
            // -----------------------------------------------------------------
            var model = new ReportModel { Name = "Aspose.Words" };

            // -----------------------------------------------------------------
            // 4. Disable reflection optimization only for this report generation.
            // -----------------------------------------------------------------
            bool originalOptimization = ReportingEngine.UseReflectionOptimization;
            ReportingEngine.UseReflectionOptimization = false;

            try
            {
                ReportingEngine engine = new ReportingEngine();
                // Build the report using the model and the root name "model".
                engine.BuildReport(doc, model, "model");
            }
            finally
            {
                // Restore the original optimization setting.
                ReportingEngine.UseReflectionOptimization = originalOptimization;
            }

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(resultPath);
        }
    }

    // Simple data model used by the template.
    public class ReportModel
    {
        public string Name { get; set; } = string.Empty;
    }
}
