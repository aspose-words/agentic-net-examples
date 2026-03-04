// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Rendering;

namespace AsposeWordsReportingPrint
{
    // Simple data model used by the report template.
    public class Employee
    {
        public string Name { get; set; }
        public string Position { get; set; }
        public decimal Salary { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOTX template that contains reporting tags.
            const string templatePath = @"C:\Templates\EmployeeReport.dotx";

            // Load the DOTX template. This uses the Document(string) constructor – the required load rule.
            Document reportDocument = new Document(templatePath);

            // Prepare a data source for the report.
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Position = "Developer", Salary = 75000m },
                new Employee { Name = "Jane Smith", Position = "Designer", Salary = 68000m },
                new Employee { Name = "Bob Johnson", Position = "Manager", Salary = 92000m }
            };

            // Build the report by merging the data source with the template.
            // The BuildReport(Document, object, string) overload is used – the required reporting rule.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDocument, employees, "Employees");

            // Optionally update fields (e.g., date fields) before printing.
            reportDocument.UpdateFields();

            // Print the generated report to the default printer.
            // This uses the Document.Print() method – the required print rule.
            reportDocument.Print();

            // If you need to specify printer settings (e.g., print a page range), use AsposeWordsPrintDocument.
            // Example: print pages 1‑2 on a specific printer.
            /*
            PrinterSettings settings = new PrinterSettings
            {
                PrinterName = "Microsoft Print to PDF",
                PrintRange = PrintRange.SomePages,
                FromPage = 1,
                ToPage = 2
            };
            AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(reportDocument);
            printDoc.PrinterSettings = settings;
            printDoc.Print();
            */
        }
    }
}
