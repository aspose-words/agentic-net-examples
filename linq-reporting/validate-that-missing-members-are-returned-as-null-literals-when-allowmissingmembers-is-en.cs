using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a builder to insert LINQ Reporting tags.
        DocumentBuilder builder = new DocumentBuilder();

        // Tag that references a missing member – should be rendered as an empty string.
        builder.Writeln("Missing member test: <<[missingObject.Name]>>");

        // Tag that iterates over a missing collection – each iteration should produce nothing.
        builder.Writeln("Missing collection iteration:");
        builder.Writeln("<<foreach [in missingObject]>>- <<[Name]>> <</foreach>>");

        // Prepare the reporting engine with the AllowMissingMembers option.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "Missed";

        // Build the report using an empty DataSet as the data source.
        // The empty DataSet ensures that no objects named "missingObject" exist.
        bool success = engine.BuildReport(builder.Document, new DataSet(), "");

        // Save the generated report.
        const string outputPath = "MissingMembersReport.docx";
        builder.Document.Save(outputPath);

        // Output a simple verification message.
        Console.WriteLine($"Report built successfully: {success}");
        Console.WriteLine($"Report saved to: {outputPath}");
    }
}
