using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReport
{
    // Sample data class that will be used in the LINQ query.
    public class Employee
    {
        public string Name { get; set; }
        public string Department { get; set; }
        public decimal Salary { get; set; }
    }

    // Example of a type that lives in another assembly.
    // In a real scenario this type would be defined in a separate project/assembly.
    public class ExternalHelper
    {
        public static string FormatSalary(decimal salary) => salary.ToString("C");
    }

    class Program
    {
        static void Main()
        {
            // Load the template DOCX that contains Aspose.Words LINQ report tags.
            // Example tag in the document: <<foreach [emp]>><<[Name]>> - <<[Salary]:formatSalary>><</foreach>>
            Document template = new Document("TemplateReport.docx");

            // Prepare a collection of data objects.
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Department = "Finance", Salary = 72000m },
                new Employee { Name = "Jane Smith", Department = "HR", Salary = 65000m },
                new Employee { Name = "Bob Johnson", Department = "IT", Salary = 85000m }
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Register the external assembly (or specific types) so that the template can call its static members.
            // Here we add the type that contains the custom formatting method.
            engine.KnownTypes.Add(typeof(ExternalHelper));

            // Build the report using the LINQ data source.
            // The engine will evaluate the template tags against the provided collection.
            engine.BuildReport(template, employees);

            // Save the generated report.
            template.Save("GeneratedReport.docx");
        }
    }
}
