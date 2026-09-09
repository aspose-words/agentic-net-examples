using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "Template.docx");
        string resultPath = Path.Combine(workDir, "Result.docx");

        // -------------------------------------------------
        // 1. Create a template document with tags that refer to a missing member.
        // -------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Tag that tries to access a missing object's property.
        builder.Writeln("<<[missingObject.First().Id]>>");

        // Foreach loop over a missing collection.
        builder.Writeln("<<foreach [in missingObject]>><<[Id]>><</foreach>>");

        // Save the template to disk (required by the lifecycle rule).
        template.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template back (simulating a real scenario).
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Configure the ReportingEngine to allow missing members
        //    and provide a custom fallback message.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "Member not found";

        // Use an empty DataSet as the data source because the template
        // does not need any real data.
        DataSet emptyData = new DataSet();

        // Build the report. The third parameter (data source name) is optional
        // when we do not reference the data source object itself in the template.
        engine.BuildReport(doc, emptyData, "");

        // -------------------------------------------------
        // 4. Save the generated report.
        // -------------------------------------------------
        doc.Save(resultPath);

        Console.WriteLine($"Report generated: {resultPath}");
    }
}
