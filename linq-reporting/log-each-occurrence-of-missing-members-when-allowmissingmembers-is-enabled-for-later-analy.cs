using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Tags that reference members which do not exist in the data source.
        builder.Writeln("Customer name: <<[MissingCustomer.Name]>>");
        builder.Writeln("<<foreach [item in MissingCollection]>>Item: <<[item.Id]>> <</foreach>>");

        // Configure the reporting engine to allow missing members.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "MISSING"
        };

        // Build the report using an empty DataSet as the data source.
        bool success = engine.BuildReport(template, new DataSet(), "");

        // Log each occurrence of the missing member placeholder.
        LogMissingMembers(template, engine.MissingMemberMessage);

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        template.Save(outputPath);
    }

    private static void LogMissingMembers(Document doc, string placeholder)
    {
        // Retrieve the full text of the document.
        string fullText = doc.GetText();

        // Split into lines for easier processing.
        string[] lines = fullText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

        int occurrence = 0;
        foreach (string line in lines)
        {
            if (line.Contains(placeholder))
            {
                occurrence++;
                Console.WriteLine($"Missing member occurrence {occurrence}: \"{line.Trim()}\"");
            }
        }

        if (occurrence == 0)
        {
            Console.WriteLine("No missing members were found in the generated report.");
        }
    }
}
