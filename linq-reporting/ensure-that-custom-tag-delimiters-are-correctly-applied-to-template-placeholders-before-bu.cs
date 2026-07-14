using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "template.docx";
            string reportPath = "report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Use default LINQ Reporting tags <<[expr]>> for the placeholders.
            builder.Writeln("Customer Report");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template back for reporting.
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // -------------------------------------------------
            // 4. Build the report using a data source.
            // -------------------------------------------------
            Person data = new Person { Name = "John Doe", Age = 30 };

            // The root object name ("person") must match the placeholder prefix.
            engine.BuildReport(loadedTemplate, data, "person");

            // -------------------------------------------------
            // 5. Save the generated report.
            // -------------------------------------------------
            loadedTemplate.Save(reportPath);
        }
    }
}
