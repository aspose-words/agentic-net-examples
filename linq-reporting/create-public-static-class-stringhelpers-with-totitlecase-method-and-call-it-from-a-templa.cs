using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Helper class with a static method that will be called from the LINQ Reporting template.
    public static class StringHelpers
    {
        // Converts the input string to title case (first letter of each word capitalized).
        public static string ToTitleCase(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            // Ensure the whole string is in lower case before applying title case.
            return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(value.ToLowerInvariant());
        }
    }

    // Simple data model used as the root object for the report.
    public class PersonModel
    {
        public string Name { get; set; } = "john doe";
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting tag that calls the static helper method.
            // The tag will output the title‑cased version of the Name property.
            builder.Writeln("<<[StringHelpers.ToTitleCase(Name)]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document for reporting.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            PersonModel model = new PersonModel();

            // -----------------------------------------------------------------
            // 4. Build the report using Aspose.Words LINQ Reporting Engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // Register the helper class so its static members can be used in the template.
            engine.KnownTypes.Add(typeof(StringHelpers));

            // Build the report. The root object name is "model" because the template
            // references members directly (e.g., Name) without a prefix.
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);

            // Optional: write a confirmation to the console.
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
