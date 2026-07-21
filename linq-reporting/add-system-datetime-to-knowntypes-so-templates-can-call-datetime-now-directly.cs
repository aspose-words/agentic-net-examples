using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the generated files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create a template document that uses a static DateTime call.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Current date and time: <<[DateTime.Now]>>");
        string templatePath = Path.Combine(workDir, "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (simulating a real‑world scenario).
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure the ReportingEngine.
        //    Add System.DateTime to KnownTypes so the template can access it.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(DateTime));

        // No data source is required for static members, but BuildReport needs an object.
        // Pass an empty anonymous object as a placeholder.
        bool success = engine.BuildReport(doc, new object(), "");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(workDir, "Report.docx");
        doc.Save(outputPath);

        // The example finishes without waiting for user input.
        // (Optional) You could verify success here, but no output is required.
    }
}
