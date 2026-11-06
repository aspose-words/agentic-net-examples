using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingTextColorExample
{
    // Simple data model with a color property.
    public class ReportModel
    {
        // Returns the name of the color to apply.
        public string ColorName { get; set; } = "Blue";

        // Sample text to display (optional, can be hard‑coded in the template).
        public string Message { get; set; } = "This text should be blue.";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a new blank document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert the LINQ Reporting tag that sets the text color dynamically.
            // The expression [model.ColorName] evaluates to "Blue".
            builder.Writeln("<<textColor [model.ColorName]>><<[model.Message]>> <</textColor>>");

            // Save the template to disk (optional, but demonstrates the full workflow).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template back (simulating a real‑world scenario where the template is read from storage).
            Document doc = new Document(templatePath);

            // Prepare the data model.
            ReportModel model = new ReportModel();

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
