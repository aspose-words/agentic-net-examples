using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model classes.
    public class ReportModel
    {
        // Collection of persons to be iterated in the template.
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string FirstName { get; set; } = "";
        public string LastName { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare sample data.
            var model = new ReportModel();
            model.Persons.Add(new Person { FirstName = "John", LastName = "Doe" });
            model.Persons.Add(new Person { FirstName = "Jane", LastName = "Smith" });
            model.Persons.Add(new Person { FirstName = "Bob", LastName = "Johnson" });

            // 2. Create a template document programmatically.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Begin a foreach loop over the Persons collection.
            builder.Writeln("<<foreach [p in Persons]>>");

            // Create a bookmark whose name is the concatenation of FirstName and LastName.
            // The expression uses string.Concat to join the two fields with a space.
            builder.Writeln("<<bookmark [string.Concat(p.FirstName, \" \", p.LastName)]>>");

            // Content inside the bookmark – just display the person's full name.
            builder.Writeln("Name: <<[p.FirstName]>> <<[p.LastName]>>");

            // Close the bookmark and the foreach block.
            builder.Writeln("<</bookmark>>");
            builder.Writeln("<</foreach>>");

            // 3. Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 4. Save the generated document.
            const string outputPath = "ReportWithBookmarks.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Report saved to {outputPath}");

            // 5. List the generated bookmark names to verify concatenation.
            Console.WriteLine("Generated bookmark names:");
            foreach (var bookmark in doc.Range.Bookmarks)
            {
                Console.WriteLine($"- {bookmark.Name}");
            }
        }
    }
}
