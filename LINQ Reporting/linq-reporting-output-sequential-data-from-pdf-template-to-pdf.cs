using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the PDF template that contains Aspose.Words reporting tags.
        Document template = new Document("Template.pdf");

        // Sample data source – a list of POCO objects.
        var people = new List<Person>
        {
            new Person { Id = 1, Name = "Alice",   Age = 30 },
            new Person { Id = 2, Name = "Bob",     Age = 25 },
            new Person { Id = 3, Name = "Charlie", Age = 35 }
        };

        // Use LINQ to order the data sequentially (by Id in this case).
        var orderedData = people.OrderBy(p => p.Id).ToList();

        // Populate the template using the ReportingEngine.
        // The third argument ("people") is the name used inside the template tags.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, orderedData, "people");

        // Save the resulting document as PDF.
        template.Save("Result.pdf", SaveFormat.Pdf);
    }

    // Simple POCO class that serves as the data source for the report.
    public class Person
    {
        public int Id   { get; set; }
        public string Name { get; set; }
        public int Age  { get; set; }
    }
}
