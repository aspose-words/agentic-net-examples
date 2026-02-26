using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    // Simple data class that will be used as the LINQ data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the DOC template that contains LINQ Reporting tags, e.g. <<foreach [persons]>><<[Name]>> (<[Age]>)<</foreach>>.
            Document template = new Document("Template.docx");

            // Prepare sequential data (could be generated via LINQ; here we use a simple list for clarity).
            List<Person> persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 45 },
                new Person { Name = "Carol", Age = 27 }
            };

            // Build the report by populating the template with the data source.
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("persons") must match the name used in the template tags.
            engine.BuildReport(template, persons, "persons");

            // Save the populated document to PostScript format using PsSaveOptions.
            PsSaveOptions psOptions = new PsSaveOptions
            {
                SaveFormat = SaveFormat.Ps // Explicitly set the format (optional, as Save method infers from options).
            };
            template.Save("ReportOutput.ps", psOptions);
        }
    }
}
