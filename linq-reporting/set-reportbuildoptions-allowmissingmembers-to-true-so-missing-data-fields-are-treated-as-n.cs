using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that references a member which does not exist.
        // The engine will treat this missing member as null because we enable AllowMissingMembers.
        builder.Writeln("<<[MissingObject.First().Id]>>");

        // Configure the reporting engine to allow missing members.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };
        // Optional: customize the message shown for a plain missing member reference.
        engine.MissingMemberMessage = "Missing";

        // Use an empty DataSet as the data source – it does not contain "MissingObject".
        DataSet dataSource = new DataSet();

        // Build the report. The third argument is the name of the data source object in the template.
        // An empty string means the template does not need to reference the data source object itself.
        engine.BuildReport(doc, dataSource, "");

        // Save the generated document.
        doc.Save("ReportAllowMissingMembers.docx");
    }
}
