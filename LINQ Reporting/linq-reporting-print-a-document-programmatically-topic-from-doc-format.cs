using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingPrint
{
    // Simple data model for the report.
    public class Employee
    {
        public string Name { get; set; }
        public string Department { get; set; }
        public decimal Salary { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOC template that contains reporting tags, e.g. <<foreach [emp]>><<[Name]>> - <<[Department]>> - $<<[Salary]>> <</foreach>>
            Document doc = new Document(@"C:\Templates\EmployeeReport.doc");

            // Prepare a collection of employees.
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Department = "Finance", Salary = 72000m },
                new Employee { Name = "Jane Smith", Department = "HR", Salary = 65000m },
                new Employee { Name = "Bob Johnson", Department = "IT", Salary = 80000m }
            };

            // Use LINQ to filter or sort the collection if needed.
            // Example: only employees with salary > 65000, ordered by salary descending.
            var filtered = employees
                .Where(e => e.Salary > 65000m)
                .OrderByDescending(e => e.Salary)
                .ToList();

            // Build the report using the ReportingEngine.
            // The data source name "emp" must match the tag used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, filtered, "emp");

            // -----------------------------------------------------------------
            // Printing
            // -----------------------------------------------------------------
            // The Document.Print method is only available on the full .NET Framework on Windows
            // and requires the Aspose.Words.Printing assembly. If you are targeting .NET Core/.NET 5+
            // or a non‑Windows platform, use an alternative approach such as saving to PDF and
            // printing via an external tool.
            // -----------------------------------------------------------------
            // Uncomment the line below if you are on .NET Framework (Windows) and have added the
            // Aspose.Words.Printing reference.
            // doc.Print();

            // Alternative for cross‑platform / .NET Core projects: save to PDF and let the OS
            // handle the printing.
            string pdfPath = @"C:\Output\EmployeeReport.pdf";
            doc.Save(pdfPath, SaveFormat.Pdf);
            // Example of launching the default PDF viewer's print command (Windows only).
            // System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(pdfPath) { Verb = "print", CreateNoWindow = true });
        }
    }
}
