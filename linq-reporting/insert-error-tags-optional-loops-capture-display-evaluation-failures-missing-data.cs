using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ------------------------------------------------------------
        // Template with an optional loop.
        // The loop iterates over a collection named "Items".
        // Inside the loop we reference two members: Name and Value.
        // If the collection or any member is missing, we want to see "<<error>>".
        // ------------------------------------------------------------
        builder.Writeln("Report start");
        builder.Writeln("<<foreach [in Items]>>");          // optional loop start
        builder.Writeln("Item: <<[Name]>> - <<[Value]>>"); // member references
        builder.Writeln("<</foreach>>");                    // optional loop end
        builder.Writeln("Report end");

        // Configure the ReportingEngine:
        // - AllowMissingMembers: missing members are treated as null instead of throwing.
        // - MissingMemberMessage: the string inserted when a plain member reference is missing.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "<<error>>"
        };

        // Build the report using an empty data source (no "Items" table, no members).
        // The engine will replace each missing member with the "<<error>>" tag.
        DataSet emptyData = new DataSet(); // no tables defined
        engine.BuildReport(doc, emptyData);

        // Save the generated document.
        doc.Save("ReportWithMissingData.docx");
    }
}
