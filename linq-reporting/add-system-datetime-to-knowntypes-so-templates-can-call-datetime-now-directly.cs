using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Empty model; required as root object for the report.
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        // Use correct syntax for static DateTime.Now.
        builder.Writeln("Current date and time: <<[DateTime.Now]>>");

        // Save the template to a local file.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template document for reporting.
        Document reportDoc = new Document(templatePath);

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // Add System.DateTime to the known types collection.
        engine.KnownTypes.Add(typeof(DateTime));

        // Build the report using an empty model as the root object.
        Model model = new Model();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine("Report generated: " + Path.GetFullPath(outputPath));
    }
}
