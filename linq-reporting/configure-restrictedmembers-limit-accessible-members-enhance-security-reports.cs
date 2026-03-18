using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a simple template document.
        // The template tries to obtain the base type of an empty string.
        // This operation would normally access System.Type members.
        // ------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("<<var [typeVar = \"\".GetType().BaseType]>><<[typeVar]>>");

        // ------------------------------------------------------------
        // 2. Restrict types whose members must not be accessible from templates.
        // Here we block System.Type, which prevents the template from calling GetType().
        // This call must be made before any report is built.
        // ------------------------------------------------------------
        ReportingEngine.SetRestrictedTypes(typeof(System.Type));

        // ------------------------------------------------------------
        // 3. Configure the ReportingEngine.
        // AllowMissingMembers prevents exceptions when a restricted member is accessed;
        // the engine will treat the missing member as a null/empty value.
        // ------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = string.Empty // optional: customize message for missing members
        };

        // ------------------------------------------------------------
        // 4. Build the report. Because System.Type is restricted,
        // the placeholder resolves to an empty string.
        // ------------------------------------------------------------
        engine.BuildReport(doc, new object());

        // ------------------------------------------------------------
        // 5. Save the resulting document.
        // ------------------------------------------------------------
        doc.Save("RestrictedReport.docx");
    }
}
