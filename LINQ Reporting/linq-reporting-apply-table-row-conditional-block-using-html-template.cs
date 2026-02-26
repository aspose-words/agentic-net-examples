using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // HTML template containing a table.
        // The <<foreach>> block iterates over the collection "persons".
        // The <<if>> block conditionally renders the row cell only when Show == true.
        string htmlTemplate = @"
<table border='1' style='border-collapse:collapse;'>
    <tr><th>Name</th></tr>
    <<foreach [in persons]>>
    <tr>
        <<if [Show]>>
        <td><<[Name]>></td>
        <<endif>>
    </tr>
    <<endforeach>>
</table>";

        // Insert the HTML template into the document.
        builder.InsertHtml(htmlTemplate);

        // Prepare the data source.
        var persons = new List<Person>
        {
            new Person { Name = "Alice",   Show = true  },
            new Person { Name = "Bob",     Show = false },
            new Person { Name = "Charlie", Show = true  }
        };

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            // Remove rows that become empty after the conditional block is evaluated.
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Build the report using the data source named "persons".
        engine.BuildReport(doc, persons, "persons");

        // Save the generated document.
        doc.Save("ConditionalTableReport.docx");
    }

    // Simple data entity used by the template.
    public class Person
    {
        public string Name { get; set; }
        public bool Show { get; set; }
    }
}
