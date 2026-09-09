using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary files.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";
        const string logPath = "ErrorLog.txt";

        // -----------------------------------------------------------------
        // 1. Create a template document with a valid tag and an invalid tag.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Valid expression – will be replaced with the model's Name property.
        builder.Writeln("Hello <<[model.Name]>>");

        // Invalid expression – missing a closing bracket, will cause a syntax error.
        builder.Writeln("This line contains an invalid tag: <<[model.Invalid>>");

        // Save the template to disk so that it can be loaded later.
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (ensures the document is fully loaded before building).
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        var model = new Model(); // Model.Name is initialized in the class definition.

        // -----------------------------------------------------------------
        // 4. Configure the ReportingEngine without InlineErrorMessages.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            // No InlineErrorMessages flag – syntax errors will throw an exception.
            Options = ReportBuildOptions.None
        };

        // -----------------------------------------------------------------
        // 5. Attempt to build the report and capture any syntax errors.
        // -----------------------------------------------------------------
        try
        {
            // The root object name used in the template tags is "model".
            bool success = engine.BuildReport(loadedTemplate, model, "model");

            // If BuildReport returns true, the report was generated without syntax errors.
            if (success)
            {
                loadedTemplate.Save(reportPath);
                Console.WriteLine($"Report generated successfully: {reportPath}");
            }
        }
        catch (Exception ex)
        {
            // Log the exception details to a file.
            File.WriteAllText(logPath, ex.ToString());
            Console.WriteLine($"An error occurred while building the report. Details logged to: {logPath}");
        }
    }
}

// ---------------------------------------------------------------------
// Simple data model used by the template.
// ---------------------------------------------------------------------
public class Model
{
    // Initialized to avoid nullable warnings.
    public string Name { get; set; } = "World";
}
