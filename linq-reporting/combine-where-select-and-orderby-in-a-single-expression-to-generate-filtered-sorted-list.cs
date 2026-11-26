using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // Sample data
            var persons = new List<Person>
            {
                new() { Name = "Alice", Age = 28 },
                new() { Name = "Bob", Age = 35 },
                new() { Name = "Charlie", Age = 32 },
                new() { Name = "Diana", Age = 24 },
                new() { Name = "Eve", Age = 40 }
            };

            // LINQ: filter Age >= 30, order by Name, select same objects
            var filteredSorted = persons
                .Where(p => p.Age >= 30)
                .OrderBy(p => p.Name)
                .Select(p => new Person { Name = p.Name, Age = p.Age })
                .ToList();

            var model = new ReportModel { Persons = filteredSorted };

            // Create template document
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Filtered and Sorted Persons:");
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
            builder.Writeln("<</foreach>>");
            doc.Save(templatePath);

            // Load template and build report
            var loadedDoc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(loadedDoc, model, "model");
            var outputPath = "Report.docx";
            loadedDoc.Save(outputPath);
        }
    }
}
