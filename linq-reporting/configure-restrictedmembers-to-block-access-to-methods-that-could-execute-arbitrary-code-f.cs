using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The tag attempts to obtain the base type of a string (System.Type) and then output it.
        // This will be blocked by the restricted type configuration.
        builder.Writeln("<<var [typeVar = \"\".GetType().BaseType]>>");
        builder.Writeln("<<[typeVar]>>");

        // Restrict access to System.Type members (and derived types) for security.
        ReportingEngine.SetRestrictedTypes(typeof(System.Type));

        // Allow missing members so the engine does not throw when the restricted member is accessed.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Build the report. The root object is not used, so we pass an empty object.
        engine.BuildReport(doc, new object());

        // Save the resulting document.
        const string outputPath = "RestrictedReport.docx";
        doc.Save(outputPath);

        // Output the document text to the console (should be empty because the access was blocked).
        Console.WriteLine("Report generated. Document text:");
        Console.WriteLine(doc.GetText().Trim());
    }
}
