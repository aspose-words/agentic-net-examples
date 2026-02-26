using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReporting
{
    // Simple data model used as the LINQ data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Load the PDF template that contains Aspose.Words reporting tags.
            //    The Document constructor handles loading; no custom creation code is required.
            Document template = new Document("Template.pdf");

            // 2. Prepare a LINQ data source.
            //    Here we create a list of Person objects and then filter it with LINQ.
            List<Person> allPeople = GetSampleData();

            // Example LINQ query: select only people older than 30 and order by name.
            var filteredPeople = allPeople
                .Where(p => p.Age > 30)
                .OrderBy(p => p.Name)
                .ToList();

            // 3. Build the report using the ReportingEngine.
            //    Use the overload that allows referencing the data source by name ("people").
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, filteredPeople, "people");

            // 4. Configure HtmlFixed save options.
            HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
            {
                // Ensure the format is explicitly set to HtmlFixed.
                SaveFormat = SaveFormat.HtmlFixed,

                // Optional: embed images as Base64 to keep a single HTML file.
                ExportEmbeddedImages = true,

                // Optional: do not show page borders in the generated HTML.
                ShowPageBorder = false
            };

            // 5. Save the populated document as fixed‑layout HTML.
            //    The Document.Save method with SaveOptions follows the required lifecycle rule.
            template.Save("Report.html", htmlOptions);
        }

        // Helper method that returns a sample collection of Person objects.
        private static List<Person> GetSampleData()
        {
            return new List<Person>
            {
                new Person { Name = "Alice Johnson", Age = 28, City = "New York" },
                new Person { Name = "Bob Smith", Age = 45, City = "London" },
                new Person { Name = "Carol White", Age = 34, City = "Sydney" },
                new Person { Name = "David Brown", Age = 52, City = "Toronto" },
                new Person { Name = "Eve Davis", Age = 31, City = "Paris" }
            };
        }
    }
}
