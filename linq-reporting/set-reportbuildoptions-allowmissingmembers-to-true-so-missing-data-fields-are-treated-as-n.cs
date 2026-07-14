using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple data model with only one property.
        var data = new SampleData { ExistingField = "Hello World" };

        // Build a template document in memory.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Write a tag that references an existing field.
        builder.Writeln("Existing: <<[data.ExistingField]>>");
        // Write a tag that references a missing field – this would normally throw an exception.
        builder.Writeln("Missing: <<[data.MissingField]>>");

        // Configure the reporting engine to treat missing members as null.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        // Optional: customize the text shown for missing members.
        engine.MissingMemberMessage = "N/A";

        // Build the report. The root object name is "data" as used in the template tags.
        engine.BuildReport(doc, data, "data");

        // Save the resulting document.
        doc.Save("ReportWithMissingMembers.docx");
    }
}

// Simple data model used as the report's data source.
public class SampleData
{
    // This property exists and will be populated in the report.
    public string ExistingField { get; set; } = string.Empty;
    // No property for MissingField – the engine will treat it as null because of the option set above.
}
