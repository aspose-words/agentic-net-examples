using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExtensionDemo
{
    // Data model used in the template.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }

        // Property used in the template to determine adulthood.
        public bool IsAdult => Age >= 18;
    }

    class Program
    {
        static void Main()
        {
            // Prepare sample data.
            var person = new Person { Name = "Alice", Age = 23 };

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Write the template with LINQ Reporting tags.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("<<if [person.IsAdult]>>Status: Adult<</if>>");
            builder.Writeln("<<if [!person.IsAdult]>>Status: Minor<</if>>");

            // Save the template to disk.
            doc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            var loadedDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report. The root object name must match the name used in the template tags.
            engine.BuildReport(loadedDoc, person, "person");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
            loadedDoc.Save(outputPath);
        }
    }
}
