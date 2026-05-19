using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class MyUtilities
{
    // Example static method that can be used in LINQ Reporting templates.
    public static string Greet(string name) => $"Hello, {name}!";
}

public class Person
{
    // Initialize to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Step 1: Create a template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert a LINQ Reporting tag that calls the static method from MyUtilities.
        // The root data source will be referenced as "person".
        builder.Writeln("<<[MyUtilities.Greet(Name)]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Register the utility class so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(MyUtilities));

        // Prepare the data source.
        var person = new Person { Name = "World" };

        // Build the report. The data source name must match the name used in the template tags.
        engine.BuildReport(reportDoc, person, "person");

        // -----------------------------------------------------------------
        // Step 3: Save the generated report.
        // -----------------------------------------------------------------
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(reportPath)}");
    }
}
