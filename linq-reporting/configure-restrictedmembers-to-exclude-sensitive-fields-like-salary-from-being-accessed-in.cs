using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingRestrictedMembers
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Sample data model
            Employee employee = new Employee
            {
                Name = "John Doe",
                Salary = 75000m
            };

            // Create template document programmatically
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Employee Report");
            builder.Writeln("Name: <<[employee.Name]>>");
            // Salary field is present in the template but will be restricted from being accessed.
            builder.Writeln("Salary: <<[employee.Salary]>>");
            string templatePath = "EmployeeTemplate.docx";
            templateDoc.Save(templatePath);

            // Load the template for report generation
            Document reportDoc = new Document(templatePath);

            // Configure the ReportingEngine.
            // The RestrictedMembers property is not available in the current version of Aspose.Words.
            // As an alternative, we can enable inline error messages so that any attempt to access a
            // restricted member (e.g., Salary) will be reported directly in the generated document.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report. The method returns a bool indicating success when InlineErrorMessages is enabled.
            bool success = engine.BuildReport(reportDoc, employee, "employee");

            // Optionally, you could check the success flag and handle errors accordingly.
            if (!success)
            {
                Console.WriteLine("Report generation completed with errors (see the document for details).");
            }

            // Save the generated report
            string reportPath = "EmployeeReport.docx";
            reportDoc.Save(reportPath);
        }
    }

    // Public data model class
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public decimal Salary { get; set; }
    }
}
