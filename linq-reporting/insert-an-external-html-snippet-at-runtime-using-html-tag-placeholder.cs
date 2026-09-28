using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // HTML snippet that will be inserted into the document.
    public string HtmlSnippet { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // Step 1: Create the template document with an <<html>> tag placeholder.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report generated with an external HTML snippet:");
        // The placeholder uses the <<html>> tag and references the model property.
        builder.Writeln("<<html [model.HtmlSnippet]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template for reporting.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Step 3: Prepare the data model containing the HTML snippet.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            HtmlSnippet = "<p style='color:blue;'>Hello <b>World</b>!</p>"
        };

        // -----------------------------------------------------------------
        // Step 4: Build the report using Aspose.Words LINQ Reporting Engine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // Step 5: Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
