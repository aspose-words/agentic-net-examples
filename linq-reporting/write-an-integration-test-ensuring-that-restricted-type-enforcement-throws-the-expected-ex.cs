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

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a template document with a variable that holds a Type object.
        var templatePath = Path.Combine(outputDir, "template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("<<var [typeVar = \"\".GetType().BaseType]>>"); // typeVar is a System.Type instance.
        builder.Writeln("<<[typeVar.FullName]>>"); // Attempt to read a member of System.Type.
        doc.Save(templatePath);

        // Load the template back (required before building the report).
        var template = new Document(templatePath);

        // Restrict the System.Type type – its members must not be accessible in the template.
        ReportingEngine.SetRestrictedTypes(typeof(System.Type));

        // Build the report and verify that an exception is thrown due to the restricted type.
        bool exceptionThrown = false;
        try
        {
            var engine = new ReportingEngine();
            // No root object is needed; the template uses only the var tag.
            engine.BuildReport(template, new object());
        }
        catch (Exception ex)
        {
            exceptionThrown = true;
            Console.WriteLine($"Caught expected exception: {ex.GetType().Name}");
        }

        Console.WriteLine($"Exception thrown as expected: {exceptionThrown}");
    }
}
