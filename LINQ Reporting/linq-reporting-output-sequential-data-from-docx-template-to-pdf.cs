using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the DOCX template that contains LINQ Reporting tags, e.g. <<foreach [persons]>><<[FirstName]>><</foreach>>
        string templatePath = "Template.docx";

        // Path where the generated PDF will be saved.
        string outputPdfPath = "Report.pdf";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Prepare a sequential data source – a list of simple POCO objects.
        List<Person> persons = new List<Person>
        {
            new Person { FirstName = "John",  LastName = "Doe",     Age = 30 },
            new Person { FirstName = "Jane",  LastName = "Smith",   Age = 25 },
            new Person { FirstName = "Bob",   LastName = "Johnson", Age = 40 }
        };

        // Build the report by populating the template with the data source.
        // The third argument ("persons") must match the name used in the template tags.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, persons, "persons");

        // Save the populated document as PDF.
        doc.Save(outputPdfPath, SaveFormat.Pdf);
    }

    // Simple data class used as the LINQ Reporting data source.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public int    Age       { get; set; }
    }
}
