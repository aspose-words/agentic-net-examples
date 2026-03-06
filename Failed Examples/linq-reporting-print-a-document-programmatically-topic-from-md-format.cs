// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using System.Linq;
using Aspose.Words;

namespace AsposeWordsLinqReporting
{
    class Program
    {
        static void Main()
        {
            // Path to the Markdown template that contains merge fields like «Name» and «Age».
            const string markdownTemplatePath = @"C:\Templates\ReportTemplate.md";

            // Load the Markdown document. The constructor automatically detects the format.
            Document document = new Document(markdownTemplatePath);

            // Sample data source – an array of anonymous objects.
            var people = new[]
            {
                new { Name = "John Doe", Age = 30 },
                new { Name = "Jane Smith", Age = 25 }
            };

            // Use LINQ to select the first record (could be any LINQ query here).
            var firstPerson = people.First();

            // Perform a simple mail‑merge to populate the template.
            document.MailMerge.Execute(
                new[] { "Name", "Age" },
                new object[] { firstPerson.Name, firstPerson.Age });

            // Ensure the layout is up‑to‑date before printing.
            document.UpdatePageLayout();

            // Print the document to the default printer.
            document.Print();

            // Optional: if you need to print to a specific printer or with custom settings,
            // create a PrinterSettings object and use the overloads of Print.
            // PrinterSettings settings = new PrinterSettings { PrinterName = "MyPrinter" };
            // document.Print(settings);
        }
    }
}
