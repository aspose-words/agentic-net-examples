using System;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Sample data
        Employee employee = new Employee
        {
            Name = "John Doe",
            Position = "Software Engineer",
            Salary = 95000m
        };

        // Create a template document programmatically
        const string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Employee Report");
        builder.Writeln("Name: <<[emp.Name]>>");
        builder.Writeln("Position: <<[emp.Position]>>");
        // Salary field will be restricted
        builder.Writeln("Salary: <<[emp.Salary]>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // Configure ReportingEngine
        ReportingEngine engine = new ReportingEngine();

        // Set RestrictedMembers via reflection (property may not exist in older versions)
        PropertyInfo? restrictedProp = typeof(ReportingEngine).GetProperty("RestrictedMembers");
        if (restrictedProp != null && restrictedProp.CanWrite)
        {
            restrictedProp.SetValue(engine, new[] { "Salary" });
        }

        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report
        bool success = engine.BuildReport(reportDoc, employee, "emp");

        // Save the generated report
        const string outputPath = "report.docx";
        reportDoc.Save(outputPath);

        // Output result information (non‑interactive)
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }
}

// Public data model
public class Employee
{
    public string Name { get; set; } = string.Empty;
    public string Position { get; set; } = string.Empty;
    public decimal Salary { get; set; }
}
