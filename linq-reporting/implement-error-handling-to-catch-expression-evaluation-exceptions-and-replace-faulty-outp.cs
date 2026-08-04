using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    public int Value { get; set; } = 10;
    public int Divisor { get; set; } = 0; // Will cause division by zero
}

public class Program
{
    public static void Main()
    {
        // 1. Create a template document with a LINQ Reporting tag that will cause an exception.
        var template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("Result: <<[model.Value / model.Divisor]>>");

        const string templatePath = "Template.docx";
        template.Save(templatePath); // Save the template to disk.

        // 2. Load the template for reporting.
        var doc = new Document(templatePath);

        // 3. Prepare the data model.
        var model = new Model();

        // 4. Configure the reporting engine to inline error messages.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // 5. Build the report. Catch any evaluation exceptions.
        bool success;
        try
        {
            success = engine.BuildReport(doc, model, "model");
        }
        catch (Exception ex)
        {
            // If an exception occurs, treat the build as failed.
            Console.WriteLine($"Report generation error: {ex.Message}");
            success = false;
        }

        // 6. Replace any inline error messages with a placeholder text.
        // Aspose.Words inserts the error message as plain text, e.g., "Error evaluating expression".
        // The regular expression removes the whole error line.
        doc.Range.Replace(new Regex(@"Error.*?(?=\r|\n|$)"), "[Invalid]");

        // 7. Save the final report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        // 8. Output simple status (no interactive input).
        Console.WriteLine($"Report generation {(success ? "succeeded" : "had errors")}. Output saved to {outputPath}");
    }
}
