using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public bool ShowSection1 { get; set; } = false;
    public bool ShowSection2 { get; set; } = false;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report Start");

        // Conditional block: will be removed entirely if both conditions are false.
        builder.Writeln("<<if [model.ShowSection1 || model.ShowSection2]>>");
        builder.Writeln("This section appears only if at least one condition is true.");
        builder.Writeln("<</if>>");

        builder.Writeln("Report End");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Data model where both booleans are false.
        ReportModel model = new()
        {
            ShowSection1 = false,
            ShowSection2 = false
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Output saved to: {Path.GetFullPath(outputPath)}");
    }
}
