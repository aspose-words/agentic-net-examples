using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a simple template document that tries to access a type's members.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("<<var [typeVar = \"\".GetType().BaseType]>><<[typeVar]>>");

        // Block access to types that can execute arbitrary code.
        // This must be done before any report is built.
        ReportingEngine.SetRestrictedTypes(
            typeof(System.Type),                 // Prevent reflection on System.Type
            typeof(System.Reflection.Assembly),  // Prevent loading assemblies
            typeof(System.Diagnostics.Process),  // Prevent starting processes
            typeof(System.IO.File),              // Prevent file system access
            typeof(System.IO.Directory)          // Prevent directory manipulation
        );

        // Configure the reporting engine.
        // AllowMissingMembers prevents exceptions when a restricted member is accessed.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Build the report using an empty data source (no actual data needed for this example).
        engine.BuildReport(doc, new object());

        // Save the resulting document.
        doc.Save("RestrictedReport.docx");
    }
}
