using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        Customer customer = new Customer
        {
            Name = "John Doe",
            Address = "123 Main St, Anytown",
            Email = "john.doe@example.com"
        };

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Customer Report");
        builder.Writeln("Name: <<[customer.Name]>>");
        builder.Writeln("Address: <<[customer.Address]>>");
        builder.Writeln("Email: <<[customer.Email]>>");

        // Save the template to disk (optional, demonstrates load‑save workflow).
        const string templatePath = "CustomerTemplate.docx";
        template.Save(templatePath);

        // Load the template back (simulating a real‑world scenario where the template exists on disk).
        Document loadedTemplate = new Document(templatePath);

        // Build the report using LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, customer, "customer");

        // Save the generated report.
        const string outputPath = "CustomerReport.docx";
        loadedTemplate.Save(outputPath);
    }
}

// Public data model used by the template.
public class Customer
{
    public string Name { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
}
