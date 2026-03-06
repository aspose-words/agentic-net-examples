using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model that will be used as the LINQ data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains Reporting Engine tags,
            // e.g. <<[ds.Name]>> and <<[ds.Age]>>.
            const string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Path where the generated report will be saved.
            const string outputPath = @"C:\Reports\GeneratedReport.docx";

            // Load the template document.
            Document templateDoc = new Document(templatePath);

            // Prepare a LINQ data source – a list of Person objects.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice Johnson", Age = 34 },
                new Person { Name = "Bob Smith",    Age = 28 },
                new Person { Name = "Carol Lee",    Age = 45 }
            };

            // The ReportingEngine can work with any .NET object.
            // We give the data source a name ("ds") so that the template can reference it.
            ReportingEngine engine = new ReportingEngine();

            // Build the report by merging the data source with the template.
            // The overload with three parameters allows us to reference the data source object itself.
            engine.BuildReport(templateDoc, people, "ds");

            // Save the populated document to the desired location.
            templateDoc.Save(outputPath);
        }
    }
}
