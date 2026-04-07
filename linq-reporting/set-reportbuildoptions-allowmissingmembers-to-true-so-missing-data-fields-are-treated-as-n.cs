using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags that reference a non‑existent data source.
        builder.Writeln("<<[missingObject.First().Id]>>");
        builder.Writeln("<<foreach [in missingObject]>><<[Id]>><</foreach>>");

        // Configure the reporting engine to treat missing members as null.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "Missed";

        // Use an empty DataSet as the data source – it does not contain 'missingObject'.
        DataSet emptyData = new DataSet();

        // Build the report. The third argument (data source name) is empty because we do not reference the root object directly.
        engine.BuildReport(doc, emptyData, "");

        // Save the generated document to a deterministic location.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Report_AllowMissingMembers.docx");
        doc.Save(outputPath);
    }
}
