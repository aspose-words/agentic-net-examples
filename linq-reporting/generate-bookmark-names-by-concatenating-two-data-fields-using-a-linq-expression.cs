using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingBookmarkExample
{
    // Data model classes
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string FirstName { get; set; } = "";
        public string LastName { get; set; } = "";
    }

    class Program
    {
        static void Main()
        {
            // Register code page provider (required for Aspose.Words)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { FirstName = "John", LastName = "Doe" },
                    new Person { FirstName = "Jane", LastName = "Smith" },
                    new Person { FirstName = "Bob", LastName = "Johnson" }
                }
            };

            // Create a template document programmatically
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert a heading
            builder.Writeln("Bookmark Example Report");
            builder.Writeln();

            // Begin foreach loop over Persons collection
            builder.Writeln("<<foreach [person in Persons]>>");

            // Create a bookmark whose name is the concatenation of FirstName and LastName
            builder.Writeln("<<bookmark [person.FirstName + \"_\" + person.LastName]>>");
            // Content inside the bookmark
            builder.Writeln("<<[person.FirstName]>> <<[person.LastName]>>");
            // Close bookmark
            builder.Writeln("<</bookmark>>");

            // End foreach loop
            builder.Writeln("<</foreach>>");

            // Save the template
            doc.Save(templatePath);

            // Load the template for report generation
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model as the root object named "model"
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report
            var outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
