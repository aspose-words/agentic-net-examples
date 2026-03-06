using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsExample
{
    // Simple data model that will be used as the data source for the template.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class SequentialReport
    {
        public static void Main()
        {
            // Path to the DOCX template that contains Aspose.Words reporting tags,
            // e.g. <<foreach [person]>><<[Name]>> (Age: <<[Age]>>)<</foreach>>.
            string templatePath = @"C:\Templates\PeopleTemplate.docx";

            // Path where the populated report will be saved.
            string outputPath = @"C:\Reports\PeopleReport.docx";

            // Create a list of Person objects – this will be the sequential data source.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 45 },
                new Person { Name = "Carol", Age = 27 }
            };

            // Load the template document (create + load lifecycle).
            Document doc = new Document(templatePath);

            // Initialize the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the list as the data source.
            // The overload with a data source name allows the template to reference the collection itself.
            // In the template the collection is referenced as "person".
            engine.BuildReport(doc, people, "person");

            // Save the populated document (save lifecycle).
            doc.Save(outputPath);
        }
    }
}
