using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // The color name or HTML color code that will be applied to the text.
        public string Color { get; set; } = "Blue";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a Word document that serves as the LINQ Reporting template.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a paragraph that uses the <<textColor>> tag.
            // The tag will evaluate the expression [model.Color] at build time.
            builder.Writeln("<<textColor [model.Color]>>This text is colored dynamically<</textColor>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (simulating a real-world scenario where the
            //    template is stored separately from the code).
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                // You can change this value to any known color name or HTML hex code.
                Color = "DarkRed"
            };

            // -----------------------------------------------------------------
            // 4. Build the report using Aspose.Words LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "model".
            engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            loadedTemplate.Save(reportPath);

            // The program finishes without waiting for user input.
        }
    }
}
