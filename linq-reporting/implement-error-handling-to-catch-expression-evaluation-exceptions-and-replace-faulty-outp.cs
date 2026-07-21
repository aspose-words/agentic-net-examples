using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model used by the template.
    public class Model
    {
        public int Number { get; set; } = 10;
        public int Zero { get; set; } = 0;
    }

    public static void Main()
    {
        // Register code page provider (required for some data sources).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // 1. Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Normal expression – should render correctly.
        builder.Writeln("Value: <<[model.Number]>>");

        // Faulty expression – division by zero will throw during evaluation.
        builder.Writeln("Faulty: <<[model.Number / model.Zero]>>");

        // Save the template (optional, just for inspection).
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // 2. Prepare the data source.
        Model model = new Model();

        // 3. Configure the ReportingEngine to inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // 4. Build the report. Catch any evaluation exceptions.
        bool success;
        try
        {
            success = engine.BuildReport(template, model, "model");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Report build failed: {ex.Message}");
            success = false;
        }

        // 5. Replace any inline error messages with a placeholder text.
        // Aspose.Words inserts the word "Error" in the message; replace it with "[Invalid]".
        template.Range.Replace("Error", "[Invalid]");

        // 6. Save the final document.
        const string outputPath = "report.docx";
        template.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated. Success flag: {success}");
    }
}
