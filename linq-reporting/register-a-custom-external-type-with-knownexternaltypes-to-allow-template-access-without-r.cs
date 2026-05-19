using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a simple template that calls a static method from a custom type.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("<<[MyHelper.GetGreeting(\"World\")]>>");

        // Save the template to disk (required before building the report).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back (simulating a real‑world scenario where the template is a file).
        Document doc = new Document(templatePath);

        // Register the custom external type so its static members can be used in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(MyHelper));

        // Build the report. No data source is needed because the template only uses static members.
        engine.BuildReport(doc, new object(), "");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Custom external type with a static method that will be accessed from the template.
public static class MyHelper
{
    public static string GetGreeting(string name) => $"Hello, {name}!";
}
