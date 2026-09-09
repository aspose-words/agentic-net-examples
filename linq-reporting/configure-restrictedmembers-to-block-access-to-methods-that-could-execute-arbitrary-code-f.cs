using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple template document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // This tag attempts to call a method that could execute arbitrary code.
        // For example, accessing System.Type.GetMethod would be dangerous.
        builder.Writeln("Attempting unsafe call: <<[typeVar.GetMethod(\"Start\")]>><<[typeVar]>>");

        // Define a variable in the template that holds a System.Type instance.
        // The variable itself is harmless, but the engine must be prevented from invoking its members.
        builder.Writeln("<<var [typeVar = typeof(System.Diagnostics.Process)]>>");

        // Restrict types that could be used to execute code.
        // Must be called before any report is built.
        ReportingEngine.SetRestrictedTypes(
            typeof(System.Type),
            typeof(System.Reflection.Assembly),
            typeof(System.Diagnostics.Process));

        // Configure the engine to allow missing members so that restricted accesses are ignored
        // and replaced with an empty string instead of throwing an exception.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };
        engine.MissingMemberMessage = string.Empty; // optional, default is empty

        // Build the report. No data source is needed for this example.
        engine.BuildReport(doc, new object());

        // Save the resulting document.
        doc.Save("RestrictedMembersExample.docx");
    }
}
