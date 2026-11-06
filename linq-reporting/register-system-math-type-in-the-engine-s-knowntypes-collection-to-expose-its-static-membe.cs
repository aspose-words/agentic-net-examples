using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags that use System.Math static members.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Value of PI: <<[Math.PI]>>");
        builder.Writeln("Sin(0.5): <<[Math.Sin(0.5)]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Dummy root model (not used in the template).
        var model = new Model();

        // Configure the reporting engine and register System.Math type.
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(Math));

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }

    // Empty model class required by the engine.
    public class Model
    {
    }
}
