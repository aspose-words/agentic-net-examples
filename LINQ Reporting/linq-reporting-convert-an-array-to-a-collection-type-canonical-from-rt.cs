using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Loading;

namespace AsposeWordsLinqReportingExample
{
    // Simple data class used as a data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }

        public Person(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the RTF template that contains LINQ Reporting tags.
            // Example template content:
            //   <<foreach [people]>>
            //   Name: <<[Name]>>, Age: <<[Age]>>
            //   <<endforeach>>
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.rtf");

            // Load the RTF document using RtfLoadOptions.
            RtfLoadOptions loadOptions = new RtfLoadOptions();
            Document doc = new Document(templatePath, loadOptions);

            // Create an array of Person objects – this is the source data.
            Person[] peopleArray = new Person[]
            {
                new Person("Alice", 30),
                new Person("Bob",   45),
                new Person("Carol", 27)
            };

            // The ReportingEngine can work directly with an array, but to demonstrate
            // conversion from a collection to an array we first obtain a collection
            // (e.g., a List) and then call ToArray().
            var peopleList = new System.Collections.Generic.List<Person>(peopleArray);
            Person[] canonicalArray = peopleList.ToArray(); // Convert collection to array.

            // Build the report using the canonical array as the data source.
            ReportingEngine engine = new ReportingEngine();
            // The third argument is the name used inside the template to reference the data source.
            engine.BuildReport(doc, canonicalArray, "people");

            // Save the populated document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
            doc.Save(outputPath);
        }
    }
}
