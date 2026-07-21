using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        string templatePath = "HyperlinkTemplate.docx";
        string outputPath = "HyperlinkReport.docx";

        // ---------- Create the template document ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a paragraph with a LINQ Reporting link tag.
        // The first expression converts the target object to a string via ToString().
        // The second expression provides the display text.
        builder.Writeln("<<link [model.Target.ToString()] [model.DisplayText]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------- Prepare the data model ----------
        ReportModel model = new ReportModel
        {
            Target = new HyperlinkTarget { Address = "https://www.example.com" },
            DisplayText = "Visit Example"
        };

        // ---------- Load the template and build the report ----------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the model. The root name must match the tag prefix ("model").
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    // The hyperlink target object. Its ToString() will be used in the link tag.
    public HyperlinkTarget Target { get; set; } = new HyperlinkTarget();

    // Text that will be displayed for the hyperlink.
    public string DisplayText { get; set; } = string.Empty;
}

// Simple class representing a hyperlink target.
// Overriding ToString() allows the object to be converted to a string in the template.
public class HyperlinkTarget
{
    public string Address { get; set; } = string.Empty;

    public override string ToString()
    {
        return Address;
    }
}
