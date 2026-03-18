using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class DisableAllowMissingMembersExample
{
    static void Main()
    {
        // Create a new document and add a template that references a missing member.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // This tag tries to access a non‑existent member "MissingObject.Id".
        builder.Writeln("<<[MissingObject.Id]>>");

        // Initialize the reporting engine without the AllowMissingMembers flag.
        // Setting Options to ReportBuildOptions.None (the default) will cause the engine
        // to throw an exception when it encounters a missing member.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None // Disable AllowMissingMembers
        };

        try
        {
            // Build the report using an empty DataSet as the data source.
            // The missing member will trigger an exception.
            engine.BuildReport(doc, new DataSet());
        }
        catch (Exception ex)
        {
            // Expected: an exception indicating that the member is missing.
            Console.WriteLine("Exception caught: " + ex.Message);
        }

        // Optionally, save the document (it will contain the original template tag).
        doc.Save("ReportWithMissingMember.docx");
    }
}
