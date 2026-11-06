using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRestrictedMembersExample
{
    // Sample data model.
    public class Employee
    {
        public string Name { get; set; } = "John Doe";
        public decimal Salary { get; set; } = 75000m; // Sensitive field.
        public string Position { get; set; } = "Software Engineer";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any encoding needs.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output.
            const string templatePath = "EmployeeTemplate.docx";
            const string outputPath = "EmployeeReport.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Employee Report");
            builder.Writeln("Name: <<[employee.Name]>>");
            builder.Writeln("Position: <<[employee.Position]>>");
            // This field is sensitive and will be restricted.
            builder.Writeln("Salary: <<[employee.Salary]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for report generation.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            var employee = new Employee();

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();

            // Enable inline error messages so attempts to access restricted members
            // (e.g., Salary) will be shown directly in the generated document.
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // -----------------------------------------------------------------
            // 5. Build the report.
            // -----------------------------------------------------------------
            bool success = engine.BuildReport(doc, employee, "employee");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(outputPath);

            // Simple console output to indicate completion.
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}.");
            Console.WriteLine($"Template: {templatePath}");
            Console.WriteLine($"Report:   {outputPath}");
        }
    }
}
