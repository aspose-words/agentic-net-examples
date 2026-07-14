using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create a template document with a tag that tries to access a
        //    restricted type (System.Type). The tag attempts to obtain the
        //    base type of an empty string.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("Attempt to access restricted type:");
        // The 'var' tag creates a variable 'typeVar' by calling GetType().
        builder.Writeln("<<var [typeVar = \"\".GetType().BaseType]>>");
        // Output the variable. If the type is restricted the result will be empty.
        builder.Writeln("Result: <<[typeVar]>>");

        // Save the template to disk.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and configure the ReportingEngine.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // Restrict the System.Type type so its members cannot be accessed from templates.
        ReportingEngine.SetRestrictedTypes(typeof(System.Type));

        var engine = new ReportingEngine
        {
            // Allow missing members to avoid exceptions when the restricted type is accessed.
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Build the report. No data source is needed for this example.
        engine.BuildReport(doc, new object(), "");

        // -----------------------------------------------------------------
        // 3. Save the generated document and display its text.
        // -----------------------------------------------------------------
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Output the resulting text to the console.
        Console.WriteLine("=== Generated Document Text ===");
        Console.WriteLine(doc.GetText().Trim());
    }
}
