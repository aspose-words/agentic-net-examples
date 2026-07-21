using System;
using System.Data;
using System.IO;
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
        // 1. Create a template document with LINQ Reporting tags that refer
        //    to members which do not exist in the data source.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Simple tag referencing a missing member.
        builder.Writeln("Missing single value: <<[missingObject.First().Id]>>");

        // Loop tag that iterates over a missing collection.
        builder.Writeln("Missing collection:");
        builder.Writeln("<<foreach [in missingObject]>>");
        builder.Writeln("  Item Id: <<[Id]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back for report generation.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure the ReportingEngine to allow missing members.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "[Missing]"
        };

        // Build the report using an empty DataSet as the data source.
        // The empty string for the data source name means we do not reference the
        // data source object itself in the template.
        engine.BuildReport(doc, new DataSet(), "");

        // -----------------------------------------------------------------
        // 4. Log each occurrence of the missing‑member placeholder.
        // -----------------------------------------------------------------
        const string placeholder = "[Missing]";
        string fullText = doc.GetText();

        int index = 0;
        int occurrence = 0;
        while ((index = fullText.IndexOf(placeholder, index, StringComparison.Ordinal)) != -1)
        {
            occurrence++;
            Console.WriteLine($"Missing member occurrence #{occurrence} at text position {index}.");
            index += placeholder.Length;
        }

        // If no occurrences were found, indicate that as well.
        if (occurrence == 0)
        {
            Console.WriteLine("No missing members were encountered.");
        }

        // -----------------------------------------------------------------
        // 5. Save the final report.
        // -----------------------------------------------------------------
        doc.Save(reportPath);
    }
}
