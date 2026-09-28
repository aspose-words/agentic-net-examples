using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingSelectDelegateExample
{
    // Simple data class.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // Root model for the report.
    public class ReportModel
    {
        // Collection of persons.
        public List<Person> Persons { get; set; } = new();

        // Delegate that formats a Person.
        public Func<Person, string> FormatPerson => p => $"{p.Name} ({p.Age})";

        // Exposes the formatted strings using the delegate.
        public IEnumerable<string> FormattedPersons => Persons.Select(FormatPerson);
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare output folder.
            string outputDir = "output";
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // Create the template document.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(outputDir, "template.docx");
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            builder.Writeln("Persons List:");
            // Use the property that already returns the formatted strings.
            builder.Writeln("<<foreach [formatted in FormattedPersons]>>");
            builder.Writeln("<<[formatted]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template and build the report.
            // -----------------------------------------------------------------
            Document doc = new(templatePath);

            // Sample data.
            ReportModel model = new()
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice",   Age = 30 },
                    new Person { Name = "Bob",     Age = 25 },
                    new Person { Name = "Charlie", Age = 35 }
                }
            };

            // Build the report.
            ReportingEngine engine = new();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            string resultPath = Path.Combine(outputDir, "report.docx");
            doc.Save(resultPath);
        }
    }
}
