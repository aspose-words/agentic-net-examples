using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CustomerInfo
{
    public string Name { get; set; } = "";
    public string Email { get; set; } = "";

    public CustomerInfo(string name, string email)
    {
        Name = name;
        Email = email;
    }
}

public class Program
{
    public static void Main()
    {
        // Enable code page support required by Aspose.Words in some environments.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data.
        var customer = new CustomerInfo("John Doe", "john.doe@example.com");

        // Create a template document.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer Report");
        builder.Writeln("Name: <<[CustomerInfo.Name]>>");
        builder.Writeln("Email: <<[CustomerInfo.Email]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine();

        // Build the report using the CustomerInfo instance as the root object named "CustomerInfo".
        engine.BuildReport(reportDoc, customer, "CustomerInfo");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
