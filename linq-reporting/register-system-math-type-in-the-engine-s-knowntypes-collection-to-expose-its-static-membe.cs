using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Define a folder for the generated files.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "output");
        Directory.CreateDirectory(outputDir);

        // Path of the template document.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        // Path of the final report.
        string resultPath = Path.Combine(outputDir, "Result.docx");

        // -----------------------------------------------------------------
        // 1. Create a template document that contains a LINQ Reporting tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        // The tag uses the static member Math.PI. After registration it will be resolved.
        builder.Writeln("Value of PI: <<[Math.PI]>>");
        // Save the template so that it can be loaded later (required by the lifecycle rule).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template document.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // Register System.Math so its static members are accessible in the template.
        engine.KnownTypes.Add(typeof(System.Math));

        // No data source is required for static members, but an object must be supplied.
        // The overload with three parameters allows us to omit a data source name.
        engine.BuildReport(doc, new object());

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(resultPath);

        // Inform the user where the files are located.
        Console.WriteLine($"Template saved to: {templatePath}");
        Console.WriteLine($"Report saved to:   {resultPath}");
    }
}
