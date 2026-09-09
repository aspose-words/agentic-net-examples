using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a simple template document with a LINQ Reporting tag.
            // -----------------------------------------------------------------
            const string templateFile = "Template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Hello <<[model.Name]>>!"); // tag will be replaced by the model data.
            templateDoc.Save(templateFile);

            // -----------------------------------------------------------------
            // 2. Load the template back from disk.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templateFile);

            // -----------------------------------------------------------------
            // 3. Prepare the data model that matches the tag in the template.
            // -----------------------------------------------------------------
            var model = new ReportModel { Name = "World" };

            // -----------------------------------------------------------------
            // 4. Build the report using ReportingEngine.
            //    The static property UseReflectionOptimization is set to false
            //    for this template.
            // -----------------------------------------------------------------
            // ReportingEngine does not implement IDisposable, so we instantiate it
            // without a using block.
            ReportingEngine.UseReflectionOptimization = false; // Disable reflection optimization.
            ReportingEngine engine = new ReportingEngine();

            // Populate the template with the model data.
            engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            loadedTemplate.Save("Report.docx");
        }
    }

    // -----------------------------------------------------------------
    // Simple data model used by the template.
    // -----------------------------------------------------------------
    public class ReportModel
    {
        // Initialise to avoid nullable warnings.
        public string Name { get; set; } = string.Empty;
    }
}
