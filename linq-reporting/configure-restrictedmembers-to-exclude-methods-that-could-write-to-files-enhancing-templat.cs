using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // This tag attempts to call a method that writes to a file.
        // The call will be blocked because System.IO.File will be added to the restricted types list.
        builder.Writeln("Attempt to write a file: <<[System.IO.File.WriteAllText(\"blocked.txt\", \"secret\")]>>");

        // Save the template so it can be re‑loaded before building the report.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Restrict the System.IO.File type – all its members become inaccessible in templates.
        ReportingEngine.SetRestrictedTypes(typeof(System.IO.File));

        // Configure the engine to treat missing members as null instead of throwing.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Build the report. The root data source is an empty object because the template does not use any data.
        engine.BuildReport(doc, new object(), "");

        // Save the resulting document. The restricted call will be omitted, leaving an empty string in its place.
        doc.Save("Report.docx");
    }
}
