using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Person
{
    public string Name { get; set; } = "John Doe";
    public int Age { get; set; } = 30;
}

public class ReportModel
{
    public Person Person { get; set; } = new Person();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel();

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert a tag that uses a static method from a core type (System.Math)
        // and a static method from a third‑party type (Newtonsoft.Json.JsonConvert).
        builder.Writeln("Square of Age: <<[Math.Pow(Person.Age, 2)]>>");
        builder.Writeln("Serialized Person JSON: <<[JsonConvert.SerializeObject(Person)]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var report = new Document(templatePath);
        var engine = new ReportingEngine();

        // Register external types so they can be used in the template.
        engine.KnownTypes.Add(typeof(System.Math));          // Core assembly type.
        engine.KnownTypes.Add(typeof(Newtonsoft.Json.JsonConvert)); // Third‑party assembly type.

        // Build the report using the model object with root name "model".
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        report.Save(reportPath);
    }
}
