// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Rendering;

namespace AsposeWordsLinqReportingPrint
{
    // Simple POCO class that will be used as a data source for the LINQ query.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Load the DOTX template document.
            // The Document constructor loads the file and determines the format automatically.
            Document doc = new Document("Template.dotx");

            // 2. Prepare a data source using LINQ.
            // In a real scenario the data could come from a database, but here we create it in‑memory.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 28 },
                new Person { Name = "Bob",   Age = 35 },
                new Person { Name = "Carol", Age = 42 }
            };

            // Example LINQ query – select only adults (age >= 30) and order by name.
            var adultPeople = people
                .Where(p => p.Age >= 30)
                .OrderBy(p => p.Name)
                .ToList();

            // 3. Populate the template with the data using ReportingEngine.
            // The data source name ("People") must match the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, adultPeople, "People");

            // 4. Print the resulting document.
            // Option A: simple print to the default printer.
            // doc.Print();

            // Option B: print with custom printer settings (e.g., specific printer, page range).
            PrinterSettings printerSettings = new PrinterSettings
            {
                // Replace with an installed printer name if needed.
                // PrinterName = "Your Printer Name",
                PrintRange = PrintRange.AllPages
            };

            // AsposeWordsPrintDocument provides a .NET PrintDocument implementation that respects Word page settings.
            AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc)
            {
                PrinterSettings = printerSettings
            };

            // Print the document without showing any UI.
            printDoc.Print();
        }
    }
}
