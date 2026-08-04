using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with a string to validate.
    public class Model
    {
        public string Input { get; set; } = "123-45-6789";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a line that outputs the raw input value.
            builder.Writeln("Input: <<[model.Input]>>");

            // Write a conditional line that uses Regex.IsMatch to validate the input.
            // The pattern checks for a US Social Security Number format: XXX-XX-XXXX.
            builder.Writeln(
                "Is SSN valid: <<if [Regex.IsMatch(model.Input, \"^\\\\d{3}-\\\\d{2}-\\\\d{4}$\")]>>Valid<</if>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document for reporting.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // Register the Regex type so that its static members can be used in template expressions.
            engine.KnownTypes.Add(typeof(Regex));

            // -----------------------------------------------------------------
            // 4. Build the report using the model as the data source.
            // -----------------------------------------------------------------
            Model model = new Model(); // Sample data; Input is already set.

            // The root object name in the template is "model".
            engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            loadedTemplate.Save(reportPath);

            // Optional: indicate completion.
            Console.WriteLine($"Report generated at: {reportPath}");
        }
    }
}
