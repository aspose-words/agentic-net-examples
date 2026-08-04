using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class EmptyModel
{
    // No members – used to demonstrate missing‑member handling.
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string resultPath = "Result.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document containing a tag that references a
        //    non‑existent member (<<[nonexistent]>>).
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<[nonexistent]>>"); // Missing member tag.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back from disk.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure the ReportingEngine to allow missing members.
        //    Missing members will be treated as null literals.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        // Leaving MissingMemberMessage empty means the engine will insert an empty string.
        engine.MissingMemberMessage = string.Empty;

        // The data source must be a non‑anonymous, visible type.
        var dataSource = new EmptyModel();

        // Build the report. The missing member tag will be replaced with an empty value.
        engine.BuildReport(reportDoc, dataSource);

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(resultPath);

        // -----------------------------------------------------------------
        // 5. Verify that the missing member was handled without error.
        //    The document text should be empty (or contain only whitespace).
        // -----------------------------------------------------------------
        string reportText = reportDoc.GetText().Trim();

        if (string.IsNullOrEmpty(reportText))
        {
            Console.WriteLine("Missing member handled successfully – output is empty as expected.");
        }
        else
        {
            Console.WriteLine("Unexpected output: '" + reportText + "'");
        }
    }
}
