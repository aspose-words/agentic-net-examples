using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains LINQ Reporting tags, e.g. <<foreach [in ds]>><<[Name]>> <<[Age]>> <</foreach>>
        Document doc = new Document("Template.docx");

        // Prepare a simple data source – a list of POCO objects.
        List<Person> data = new List<Person>
        {
            new Person { Name = "John Doe", Age = 30 },
            new Person { Name = "Jane Smith", Age = 25 },
            new Person { Name = "Bob Johnson", Age = 45 }
        };

        // Create the reporting engine and populate the template.
        ReportingEngine engine = new ReportingEngine();
        // The third argument ("ds") is the name used in the template to reference the data source.
        engine.BuildReport(doc, data, "ds");

        // Save the resulting document as PDF.
        doc.Save("Report.pdf", SaveFormat.Pdf);
    }

    // Simple POCO class used as a data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
