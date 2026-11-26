using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace Aspose.Words.Reporting
{
    // Minimal stub to allow registration of external types when the real class is unavailable.
    public static class KnownExternalTypes
    {
        public static void Add(Type type)
        {
            // In a full Aspose.Words environment this would register the type.
            // For this standalone example the method is intentionally left empty.
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Register the external type so that the reporting engine can access it without reflection.
        KnownExternalTypes.Add(typeof(MyExternalClass));

        // Create a simple template with a LINQ Reporting tag that accesses the external type.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Hello, <<[model.External.Greeting]>>!");
        template.Save(templatePath);

        // Create the model that contains an instance of the external type.
        ReportModel model = new ReportModel();

        // Build the report.
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        report.Save(outputPath);
    }
}

// The model passed to the reporting engine.
public class ReportModel
{
    public MyExternalClass External { get; set; } = new MyExternalClass();
}

// A custom external type that we want to expose to the template without reflection.
public class MyExternalClass
{
    public string Greeting => "World from an external type";
}
