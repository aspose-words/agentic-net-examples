using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "template.docx");
        string resultPath = Path.Combine(outputDir, "result.docx");

        // Create a template document with a tag that accesses a member of System.Type.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        // The tag attempts to retrieve the BaseType of an empty string's type.
        builder.Writeln("<<var [typeVar = \"\".GetType().BaseType]>><<[typeVar]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Restrict access to System.Type members (e.g., BaseType) for security.
        ReportingEngine.SetRestrictedTypes(typeof(System.Type));

        // Configure the reporting engine to allow missing members without throwing.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Build the report. The root data source is an empty object because the template does not use it.
        engine.BuildReport(doc, new object());

        // Save the generated report.
        doc.Save(resultPath);

        // Indicate completion.
        Console.WriteLine($"Report generated at: {resultPath}");
    }
}
