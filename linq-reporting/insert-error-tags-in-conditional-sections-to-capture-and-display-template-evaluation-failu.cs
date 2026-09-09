using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting; // ReportingEngine and ReportBuildOptions are defined here.
using System.Text;

namespace LinqReportingInlineErrorDemo
{
    // Simple data model used by the template.
    public class Model
    {
        // Controls whether the conditional block is evaluated.
        public bool ShowValue { get; set; } = true;

        // Displayed when the condition is true.
        public string Value { get; set; } = "12345";

        // Always displayed.
        public string Always { get; set; } = "Always present";

        // No property named Missing – accessing it will cause a runtime error,
        // which will be captured by the InlineErrorMessages option.
    }

    public class Program
    {
        public static void Main()
        {
            // Register the code page provider (required for some data sources).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Conditional block: will be evaluated because ShowValue == true.
            builder.Writeln("<<if [model.ShowValue]>>");
            builder.Writeln("Value: <<[model.Value]>>");
            // This line references a non‑existent member and will trigger an error.
            builder.Writeln("Missing: <<[model.Missing]>>");
            builder.Writeln("<</if>>");

            // This line is always evaluated and should render correctly.
            builder.Writeln("Always: <<[model.Always]>>");

            // Save the template to a local file.
            const string templatePath = "template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.InlineErrorMessages
            };

            // Build the report using the model as the data source.
            Model model = new Model();
            bool success = engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "report.docx";
            doc.Save(outputPath);

            // Inform the user about the result.
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}.");
            Console.WriteLine($"Template: {Path.GetFullPath(templatePath)}");
            Console.WriteLine($"Output : {Path.GetFullPath(outputPath)}");
        }
    }
}
