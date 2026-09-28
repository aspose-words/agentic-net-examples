using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Holds the HTML snippet read from an external file.
    public string HtmlSnippet { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Ensure the working directory is writable.
        string workingDir = Directory.GetCurrentDirectory();

        // 1. Create a sample HTML snippet file.
        string htmlFilePath = Path.Combine(workingDir, "snippet.html");
        File.WriteAllText(htmlFilePath,
            "<p style=\"color:blue;\">This is <b>HTML</b> snippet inserted via LINQ Reporting.</p>");

        // 2. Build the template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report with external HTML snippet:");
        // Insert the HTML switch tag that will be replaced by the snippet content.
        builder.Writeln("<<[model.HtmlSnippet] -html>>");
        // Save the template to disk.
        string templatePath = Path.Combine(workingDir, "template.docx");
        templateDoc.Save(templatePath);

        // 3. Load the template for reporting.
        var doc = new Document(templatePath);

        // 4. Prepare the data model, reading the HTML snippet from the file.
        var model = new ReportModel
        {
            HtmlSnippet = File.ReadAllText(htmlFilePath)
        };

        // 5. Build the report using Aspose.Words LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 6. Save the generated report.
        string outputPath = Path.Combine(workingDir, "output.docx");
        doc.Save(outputPath);
    }
}
