using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a template document that references a collection which may be missing.
        DocumentBuilder builder = new DocumentBuilder();
        // The foreach block will iterate over "missingCollection". If the collection does not exist,
        // we want the engine to treat it as an empty collection rather than throwing an exception.
        builder.Writeln("<<foreach [in missingCollection]>><<[Name]>><</foreach>>");

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            // Enable handling of missing members (including missing collections).
            Options = ReportBuildOptions.AllowMissingMembers,
            // When a plain reference to a missing member is encountered, output an empty string.
            // The default is already an empty string, but we set it explicitly for clarity.
            MissingMemberMessage = string.Empty
        };

        // Use an empty DataSet as the data source – it does not contain "missingCollection".
        DataSet emptyDataSource = new DataSet();

        // Build the report. Missing collection items are treated as empty, so no null reference errors occur.
        engine.BuildReport(builder.Document, emptyDataSource, "");

        // Save the generated document.
        builder.Document.Save("Report.docx");
    }
}
