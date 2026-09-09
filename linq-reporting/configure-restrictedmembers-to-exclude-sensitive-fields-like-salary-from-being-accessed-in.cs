using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple data model with a sensitive field.
        var employee = new Employee
        {
            Name = "John Doe",
            Salary = 85000
        };

        // -----------------------------------------------------------------
        // 1. Build the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert tags that reference both a safe and a sensitive member.
        builder.Writeln("Employee Name: <<[emp.Name]>>");
        builder.Writeln("Employee Salary: <<[emp.Salary]>>"); // This should be blocked.

        // Save the template to a local file (required by the lifecycle rules).
        const string templatePath = "EmployeeReportTemplate.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (simulating a real‑world scenario).
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure the ReportingEngine.
        // -----------------------------------------------------------------
        // Restrict the Employee type so that its members cannot be accessed
        // from the template. This effectively hides the Salary field.
        ReportingEngine.SetRestrictedTypes(typeof(Employee));

        var engine = new ReportingEngine
        {
            // Allow missing members to avoid exceptions when a blocked member is used.
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // -----------------------------------------------------------------
        // 4. Build the report.
        // -----------------------------------------------------------------
        // The root object name used in the template is "emp".
        engine.BuildReport(doc, employee, "emp");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "EmployeeReport.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }

    // -----------------------------------------------------------------
    // Data model.
    // -----------------------------------------------------------------
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public decimal Salary { get; set; }
    }
}
