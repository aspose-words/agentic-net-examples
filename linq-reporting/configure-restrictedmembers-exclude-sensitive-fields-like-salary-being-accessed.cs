using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportingEngineRestrictedExample
{
    // Sample data class that contains a sensitive field.
    public class Employee
    {
        public string Name { get; set; }
        public decimal Salary { get; set; }   // Sensitive information.
    }

    // Wrapper class used as a visible data source for the reporting engine.
    public class ReportData
    {
        public Employee employee { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a simple template document with placeholders.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Employee Name: <<[employee.Name]>>");
            builder.Writeln("Employee Salary: <<[employee.Salary]>>"); // Should be blocked.

            // 2. Restrict the Employee type so that its members cannot be accessed via the reporting engine.
            ReportingEngine.SetRestrictedTypes(typeof(Employee));

            // 3. Prepare the data source.
            var employee = new Employee { Name = "John Doe", Salary = 12345.67m };
            var data = new ReportData { employee = employee };

            // 4. Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers
            };
            engine.MissingMemberMessage = string.Empty;

            // 5. Build the report.
            engine.BuildReport(doc, data);

            // 6. Output the result to the console.
            Console.WriteLine("Generated Document Text:");
            Console.WriteLine(doc.GetText().Trim());

            // 7. Save the document if needed.
            doc.Save("RestrictedReport.docx");
        }
    }
}
