using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model used as the root object for the LINQ Reporting engine.
    public class Person
    {
        // Initialize properties to avoid nullable warnings.
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a Word template programmatically and insert LINQ tags.
            // -----------------------------------------------------------------
            const string templatePath = "Template.docx";

            // Create a blank document and a builder to add content.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Simple heading.
            builder.Writeln("Report with a dynamically generated bookmark:");

            // Bookmark tag where the bookmark name is the concatenation of FirstName and LastName.
            // The expression uses the root object name "person".
            builder.Writeln("<<bookmark [person.FirstName + \"_\" + person.LastName]>>");
            // Content that will be placed inside the bookmark.
            builder.Writeln("<<[person.Title]>>");
            // Close the bookmark tag.
            builder.Writeln("<</bookmark>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Sample data for the report.
            Person person = new Person
            {
                FirstName = "John",
                LastName = "Doe",
                Title = "Senior Manager"
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "person".
            engine.BuildReport(loadedTemplate, person, "person");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            loadedTemplate.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
