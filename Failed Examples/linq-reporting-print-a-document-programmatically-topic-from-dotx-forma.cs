// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Load the DOTX template file.
        Document template = new Document("Template.dotx");

        // Prepare a LINQ‑based data source.
        // Here we filter a collection of Person objects and project an anonymous type.
        var data = GetPersons()
            .Where(p => p.Age >= 18)               // LINQ filter
            .Select(p => new { FullName = p.Name, Age = p.Age })
            .ToArray();                             // ReportingEngine expects an array

        // Populate the template with the data using ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, data, "persons");

        // Print the resulting document.
        AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(template);

        // Optional: configure printer settings (e.g., print only the first two pages).
        PrinterSettings settings = new PrinterSettings
        {
            PrintRange = PrintRange.SomePages,
            FromPage = 1,
            ToPage = 2
        };
        printDoc.PrinterSettings = settings;

        // Send the document to the default printer.
        printDoc.Print();
    }

    // Sample data provider.
    static IEnumerable<Person> GetPersons()
    {
        return new List<Person>
        {
            new Person { Name = "Alice",   Age = 25 },
            new Person { Name = "Bob",     Age = 17 },
            new Person { Name = "Charlie", Age = 32 }
        };
    }

    // Simple POCO used for the LINQ query.
    class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
