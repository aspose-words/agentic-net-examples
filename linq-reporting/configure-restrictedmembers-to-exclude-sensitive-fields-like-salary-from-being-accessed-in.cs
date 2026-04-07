using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple template document with LINQ Reporting tags.
        const string templateFile = "Template.docx";
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Employee Report");
        builder.Writeln("Name: <<[employee.Name]>>");
        builder.Writeln("Salary: <<[employee.Salary]>>");
        template.Save(templateFile);

        // Load the template for reporting.
        Document doc = new Document(templateFile);

        // Restrict the Employee type so its members cannot be accessed in the template.
        // This must be done before the first BuildReport call.
        ReportingEngine.SetRestrictedTypes(typeof(Employee));

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            // Allow missing members to avoid exceptions when a restricted member is referenced.
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = string.Empty
        };

        // Sample data source.
        Employee employee = new Employee
        {
            Name = "John Doe",
            Salary = 12345.67m
        };

        // Build the report. The Salary tag will be ignored because the Employee type is restricted.
        engine.BuildReport(doc, employee, "employee");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Data model used in the example.
public class Employee
{
    public string Name { get; set; } = string.Empty;
    public decimal Salary { get; set; }
}
