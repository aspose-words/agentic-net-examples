using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // The document to be included. It is intentionally left null to simulate a missing file.
    public Document? MissingDoc { get; set; } = null;
}

public class Program
{
    public static void Main()
    {
        // Create a working directory for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create a template document that contains a <<doc>> tag
        //    referencing a missing document (null reference).
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report start");
        // The <<doc>> tag expects a Document object. The variable name must be a valid identifier.
        builder.Writeln("<<doc [MissingDoc]>>");
        builder.Writeln("Report end");

        // Save the template so it can be loaded later.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure the ReportingEngine to treat missing members as optional.
        //    The AllowMissingMembers flag causes the engine to replace missing
        //    members (including a null Document) with null and skip them.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        // An empty message ensures nothing is inserted for the missing member.
        engine.MissingMemberMessage = string.Empty;

        // Build the report using a model that has MissingDoc set to null.
        ReportModel model = new ReportModel();
        engine.BuildReport(loadedDoc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(workDir, "ReportResult.docx");
        loadedDoc.Save(outputPath);

        Console.WriteLine("Report generated successfully: " + outputPath);
    }
}
