using System;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document that references missing members.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // The tag <<[missingObject.Name]>> refers to a member that does not exist.
        builder.Writeln("Customer: <<[missingObject.Name]>>");

        // A foreach loop over a missing collection.
        builder.Writeln("<<foreach [item in missingObject]>>Item: <<[item]>> <</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for reporting.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure the ReportingEngine to allow missing members.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "[Missing]"
        };

        // Build the report using an empty DataSet as the data source.
        // The empty string for the data source name allows direct member access.
        engine.BuildReport(reportDoc, new DataSet(), "");

        // -----------------------------------------------------------------
        // 4. Log each occurrence of the missing member placeholder.
        // -----------------------------------------------------------------
        string documentText = reportDoc.GetText();
        int missingCount = Regex.Matches(documentText, Regex.Escape("[Missing]")).Count;

        Console.WriteLine($"Missing members logged: {missingCount}");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(reportPath);
    }
}
