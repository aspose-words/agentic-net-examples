using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Sample property used in the template.
    public string CustomerName { get; set; } = "John Doe";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the temporary template and final output.
        const string templatePath = "template.docx";
        const string outputPath = "output.pdf";

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a simple paragraph with a LINQ Reporting tag.
        builder.Writeln("Report for <<[model.CustomerName]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template and build the report.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Prepare the root data object.
        ReportModel model = new ReportModel();

        // Create and configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // No special options needed.

        // Build the report using the model and the root name "model".
        engine.BuildReport(reportDoc, model, "model");

        // -------------------------------------------------
        // Write the generated report directly to an HTTP response stream.
        // In this console example we simulate the response stream with a MemoryStream.
        // -------------------------------------------------
        using (MemoryStream httpResponseStream = new MemoryStream())
        {
            // Save the report to the simulated HTTP response stream (PDF format).
            reportDoc.Save(httpResponseStream, SaveFormat.Pdf);

            // Reset the stream position so it can be read from the beginning.
            httpResponseStream.Position = 0;

            // For demonstration purposes, also write the stream contents to a file.
            File.WriteAllBytes(outputPath, httpResponseStream.ToArray());
        }

        // Clean up temporary template file.
        if (File.Exists(templatePath))
            File.Delete(templatePath);
    }
}
