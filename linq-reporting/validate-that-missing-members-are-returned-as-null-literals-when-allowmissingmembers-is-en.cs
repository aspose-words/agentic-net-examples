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

        // Insert LINQ Reporting tags that reference a missing member.
        // The engine will treat these missing members as null literals.
        builder.Writeln("Missing member test: <<[missingObject.First().Id]>>");
        builder.Writeln("Missing collection test: <<foreach [in missingObject]>><<[Id]>><</foreach>>");

        // Configure the reporting engine to allow missing members.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        // The message printed for a missing member; using the literal "null" for verification.
        engine.MissingMemberMessage = "null";

        // Build the report using an empty DataSet as the data source.
        // The root name is an empty string because we are not referencing the data source object itself.
        bool success = engine.BuildReport(doc, new DataSet(), "");

        // Output the result to the console for verification.
        Console.WriteLine("Report build successful: " + success);
        Console.WriteLine("Document content:");
        Console.WriteLine(doc.GetText());

        // Save the generated document (optional, helps visual verification).
        doc.Save("MissingMembersReport.docx");
    }
}
