using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data entity.
    public class Customer
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare a collection of customers.
            List<Customer> customers = new List<Customer>
            {
                new Customer { Name = "Thomas Hardy", Age = 45, City = "London" },
                new Customer { Name = "Paolo Accorti", Age = 38, City = "Torino" },
                new Customer { Name = "John Doe", Age = 29, City = "New York" }
            };

            // 2. Use LINQ to project the data we need for the report.
            //    Here we select only Name and City, and calculate a new field.
            var reportData = customers
                .Select(c => new
                {
                    FullName = c.Name,
                    Location = c.City,
                    AgeGroup = c.Age >= 40 ? "Senior" : "Junior"
                })
                .ToList();

            // 3. Load the Word template that contains merge tags like <<[ds.FullName]>> etc.
            //    The template file must exist on disk; adjust the path as needed.
            string templatePath = @"C:\Templates\CustomerReportTemplate.docx";
            Document doc = new Document(templatePath); // lifecycle: load

            // 4. Build the report using Aspose.Words ReportingEngine.
            //    The data source name "ds" must match the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, reportData, "ds"); // feature: BuildReport

            // 5. Save the populated document.
            string outputPath = @"C:\Output\CustomerReport.docx";
            doc.Save(outputPath); // lifecycle: save
        }
    }
}
