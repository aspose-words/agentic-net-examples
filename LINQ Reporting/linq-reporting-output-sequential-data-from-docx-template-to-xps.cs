using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingToXps
{
    // Simple data class used as the LINQ data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains LINQ Reporting tags.
            // Example template content:
            // <<foreach [items]>>
            //   Name: <<[Name]>>
            //   Age: <<[Age]>>
            // <<endforeach>>
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare a sequential data source using LINQ.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 28 },
                new Person { Name = "Bob",   Age = 35 },
                new Person { Name = "Carol", Age = 42 }
            };

            // Build the report by populating the template with the data source.
            // The data source name ("items") must match the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, people, "items");

            // Save the populated document as XPS.
            string outputPath = @"C:\Output\ReportOutput.xps";
            XpsSaveOptions saveOptions = new XpsSaveOptions(); // default options
            doc.Save(outputPath, saveOptions);
        }
    }
}
