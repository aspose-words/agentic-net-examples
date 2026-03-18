using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a template expression that will cause a syntax error.
        // The expression references a missing object, which will trigger an inline error message.
        builder.Writeln("<<[MissingObject.Id]>>");

        // Configure the reporting engine:
        // - InlineErrorMessages: embed syntax error messages directly into the output document.
        // - AllowMissingMembers: treat missing members as null instead of throwing an exception.
        // - MissingMemberMessage: custom text that will be shown for missing members.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages | ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "[Missing data]"
        };

        // Build the report using an empty data source (no tables, no objects).
        DataSet emptyData = new DataSet();
        engine.BuildReport(doc, emptyData);

        // Save the document with the inline error message.
        doc.Save("InlineErrorMessage.docx");
    }
}
