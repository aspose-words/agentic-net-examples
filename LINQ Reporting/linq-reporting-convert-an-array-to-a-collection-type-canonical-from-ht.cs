using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // HTML template that uses LINQ Reporting Engine syntax.
        // The <<foreach [items]>> tag expects a collection named "items".
        string htmlTemplate = @"
        <html><body>
        <<foreach [items]>>
        <p>Name: <<[Name]>></p>
        <<endforeach>>
        </body></html>";

        // Load the HTML string into an Aspose.Words Document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertHtml(htmlTemplate);

        // Create an array of data objects.
        Person[] peopleArray = new[]
        {
            new Person { Name = "Alice" },
            new Person { Name = "Bob" },
            new Person { Name = "Charlie" }
        };

        // Convert the array to a canonical collection type (List<T>) that the ReportingEngine can iterate.
        List<Person> peopleList = new List<Person>(peopleArray);

        // Build the report using the list as the data source.
        ReportingEngine engine = new ReportingEngine();
        // The third argument is the name used in the template to reference the collection.
        engine.BuildReport(doc, peopleList, "items");

        // Save the populated document.
        doc.Save("ReportFromHtml.docx");
    }

    // Simple data class used in the report.
    public class Person
    {
        public string Name { get; set; }
    }
}
