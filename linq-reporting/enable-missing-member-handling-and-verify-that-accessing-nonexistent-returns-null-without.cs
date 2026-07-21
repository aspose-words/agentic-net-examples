using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "output.docx");

        // 1. Create a template document with a missing-member tag.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        // The tag references a member that does not exist in the data source.
        builder.Writeln("<<[nonexistent]>>");
        templateDoc.Save(templatePath);

        // 2. Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // 3. Configure the reporting engine to allow missing members.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            // Optional: customize the message printed for missing members.
            MissingMemberMessage = string.Empty
        };

        // 4. Build the report using an empty data source.
        // The data source does not contain a 'nonexistent' member.
        object emptyDataSource = new object();
        engine.BuildReport(reportDoc, emptyDataSource, "");

        // 5. Verify that the missing member was treated as null (empty output).
        string resultText = reportDoc.GetText().Trim();
        bool isEmpty = string.IsNullOrEmpty(resultText);
        Console.WriteLine(isEmpty
            ? "Missing member handled correctly: output is empty."
            : $"Unexpected output: \"{resultText}\"");

        // 6. Save the generated report.
        reportDoc.Save(outputPath);
    }
}
