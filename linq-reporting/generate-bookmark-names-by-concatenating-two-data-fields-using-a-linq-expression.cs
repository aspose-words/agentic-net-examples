using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model classes
    public class Person
    {
        // Initialize to avoid nullable warnings
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
    }

    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { FirstName = "John", LastName = "Doe" },
                    new Person { FirstName = "Jane", LastName = "Smith" }
                }
            };

            // Create a template document programmatically
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Begin a foreach loop over the Persons collection
            builder.Writeln("<<foreach [p in Persons]>>");

            // Create a bookmark whose name is the concatenation of FirstName and LastName
            builder.Writeln("<<bookmark [p.FirstName + \"_\" + p.LastName]>>");

            // Content that will be placed inside the bookmark
            builder.Writeln("<<[p.FirstName]>> <<[p.LastName]>>");

            // Close the bookmark and the foreach block
            builder.Writeln("<</bookmark>>");
            builder.Writeln("<</foreach>>");

            // Build the report using the LINQ Reporting engine
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model);

            // Save the resulting document
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ReportWithBookmarks.docx");
            doc.Save(outputPath);

            // Inform the user (no interactive input required)
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
