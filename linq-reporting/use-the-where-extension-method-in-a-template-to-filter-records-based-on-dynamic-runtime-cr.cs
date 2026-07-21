using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingWhereExample
{
    // Simple data entity.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // Wrapper model that holds the collection and the runtime filter criteria.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
        public int MinAge { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel
            {
                MinAge = 30,
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 25 },
                    new Person { Name = "Bob",   Age = 35 },
                    new Person { Name = "Carol", Age = 30 },
                    new Person { Name = "Dave",  Age = 45 }
                }
            };

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Title showing the dynamic filter value.
            builder.Writeln("Persons with Age >= <<[data.MinAge]>>:");

            // Foreach loop over all persons.
            builder.Writeln("<<foreach [p in data.Persons]>>");
            // Conditional output – only show persons that satisfy the runtime criteria.
            builder.Writeln("<<if [p.Age >= data.MinAge]>>- <<[p.Name]>> (Age: <<[p.Age]>>)<</if>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            var loadedDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model; the root name in the template is "data".
            engine.BuildReport(loadedDoc, model, "data");

            // Save the generated report.
            const string reportPath = "Report.docx";
            loadedDoc.Save(reportPath);
        }
    }
}
