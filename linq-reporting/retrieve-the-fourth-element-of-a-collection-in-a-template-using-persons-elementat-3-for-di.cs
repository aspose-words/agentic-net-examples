using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with a collection of Person objects.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    // Person class used in the collection.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }

        public Person(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a LINQ Reporting tag that retrieves the fourth element (index 3) from the collection.
            // The root data source will be referenced as "model" in BuildReport.
            builder.Writeln("Fourth person: <<[model.Persons.ElementAt(3).Name]>> (Age: <<[model.Persons.ElementAt(3).Age]>>)");

            // Save the template to disk (required before building the report).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Step 2: Prepare sample data.
            ReportModel model = new ReportModel();
            model.Persons.Add(new Person("Alice", 30));
            model.Persons.Add(new Person("Bob", 25));
            model.Persons.Add(new Person("Charlie", 28));
            model.Persons.Add(new Person("Diana", 32)); // Fourth element (index 3)
            model.Persons.Add(new Person("Ethan", 27));

            // Step 3: Load the template (could reuse the same Document instance, but following lifecycle rules we load it).
            Document doc = new Document(templatePath);

            // Step 4: Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Step 5: Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
