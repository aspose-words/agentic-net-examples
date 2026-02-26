using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting Engine template tags.
        // The <<foreach>> tag will iterate over the "persons" collection.
        builder.Writeln("Report of persons:");
        builder.Writeln("<<foreach [persons]>><<[Name]>> - <<[Age]>>\n<</foreach>>");

        // Prepare a simple data source.
        List<Person> persons = new List<Person>
        {
            new Person { Name = "Alice",   Age = 30 },
            new Person { Name = "Bob",     Age = 45 },
            new Person { Name = "Charlie", Age = 25 }
        };

        // Build the report by populating the template with the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, persons, "persons");

        // Save the populated document as PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions(); // default options
        doc.Save("Report.pdf", pdfOptions);
    }

    // Simple POCO class used as the data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
