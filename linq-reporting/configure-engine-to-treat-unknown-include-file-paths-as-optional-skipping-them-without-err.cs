using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple wrapper model used as the data source for the report.
    public class ReportModel
    {
        // The document to be included. Leaving it null simulates a missing file.
        public Document? IncludeDoc { get; set; } = null;
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document that attempts to include an external
        //    document via the supported <<doc>> tag. The tag references the
        //    IncludeDoc property of the data source.
        // -----------------------------------------------------------------
        var builder = new DocumentBuilder();
        builder.Writeln("Document start");
        // The <<doc>> tag will try to insert the document referenced by the data source.
        // If the property is null, the engine will treat it as a missing member.
        builder.Writeln("<<doc [model.IncludeDoc]>>");
        builder.Writeln("Document end");

        const string templatePath = "Template.docx";
        builder.Document.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for reporting.
        // -----------------------------------------------------------------
        var templateDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure the reporting engine to allow missing members.
        //    This makes the engine treat a null IncludeDoc as an optional include.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        // Optional: customize the message shown for missing members.
        engine.MissingMemberMessage = string.Empty;

        // -----------------------------------------------------------------
        // 4. Build the report using a model instance where IncludeDoc is null.
        // -----------------------------------------------------------------
        var model = new ReportModel(); // IncludeDoc remains null.
        engine.BuildReport(templateDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        const string resultPath = "Result.docx";
        templateDoc.Save(resultPath);

        Console.WriteLine($"Report generated: {resultPath}");
    }
}
