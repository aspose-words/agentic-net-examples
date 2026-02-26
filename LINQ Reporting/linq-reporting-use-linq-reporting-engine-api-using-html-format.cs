using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting template directly as HTML.
        // The template iterates over a collection named "persons" and outputs each person's Name and Age.
        string template = @"<<foreach [persons]>><<[Name]>> - <<[Age]>> years<</foreach>>";
        builder.InsertHtml(template);

        // Prepare the data source: a list of Person objects.
        var persons = new List<Person>
        {
            new Person { Name = "John Doe", Age = 30 },
            new Person { Name = "Jane Smith", Age = 25 }
        };

        // Build the report using the ReportingEngine.
        // The third argument is the name used in the template to reference the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, persons, "persons");

        // Save the populated document as HTML.
        doc.Save("Report.html", SaveFormat.Html);
    }

    // Simple data class used as the report's data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
