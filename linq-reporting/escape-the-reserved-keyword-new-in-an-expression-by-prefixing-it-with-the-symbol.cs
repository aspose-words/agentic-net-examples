using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Property named "new" escaped with @ in C# code.
    public string @new { get; set; } = "Escaped keyword value";
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var model = new Model();

        // Create a template document programmatically.
        var templatePath = "template.docx";
        var builder = new DocumentBuilder();

        builder.Writeln("Value of escaped keyword property:");
        // In LINQ Reporting expressions the property name can be used directly.
        builder.Writeln("<<[model.new]>>");

        builder.Document.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
