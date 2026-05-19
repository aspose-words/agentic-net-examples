using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ExternalModels
{
    public class ModelA
    {
        public string Name => "ModelA Instance";

        public static string GetInfo()
        {
            return "Static info from ModelA";
        }
    }
}

namespace OtherModels
{
    public class ModelB
    {
        public int Value => 42;

        public static string GetDetail()
        {
            return "Static detail from ModelB";
        }
    }
}

public class ReportData
{
    public string Title { get; set; } = string.Empty;
    public ExternalModels.ModelA A { get; set; } = new();
    public OtherModels.ModelB B { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report Title: <<[model.Title]>>");
        builder.Writeln("A Name: <<[model.A.Name]>>");
        builder.Writeln("A Info (static): <<[ExternalModels.ModelA.GetInfo()]>>");
        builder.Writeln("B Value: <<[model.B.Value]>>");
        builder.Writeln("B Detail (static): <<[OtherModels.ModelB.GetDetail()]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Register external types from different namespaces.
        var engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(ExternalModels.ModelA));
        engine.KnownTypes.Add(typeof(OtherModels.ModelB));

        // Prepare the root data object.
        var data = new ReportData
        {
            Title = "LINQ Reporting Demo"
        };

        // Build the report.
        engine.BuildReport(doc, data, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "report.docx");
        doc.Save(reportPath);
    }
}
