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
            // 1. Create a template document with a placeholder for the title.
            // -----------------------------------------------------------------
            const string templateFile = "Template.docx";

            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Set the paragraph style to Heading1 before writing the placeholder.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("<<[model.Title]>>");

            // Save the template to disk.
            templateDoc.Save(templateFile);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data model.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templateFile);

            ReportModel model = new ReportModel
            {
                Title = "Report Title Using Heading1 Style"
            };

            // -----------------------------------------------------------------
            // 3. Build the report using Aspose.Words LINQ Reporting Engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputFile = "Report.docx";
            loadedTemplate.Save(outputFile);
        }
    }

    // Simple data model with a single property referenced in the template.
    public class ReportModel
    {
        public string Title { get; set; } = string.Empty;
    }
}
