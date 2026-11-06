using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Working directory.
        string workDir = Directory.GetCurrentDirectory();

        // 1. Create an external DOCX file that will be inserted.
        string externalPath = Path.Combine(workDir, "external.docx");
        Document externalDoc = new Document();
        DocumentBuilder extBuilder = new DocumentBuilder(externalDoc);
        extBuilder.Writeln("This is the content of the external document.");
        externalDoc.Save(externalPath);

        // 2. Create the template DOCX with a bookmark that includes the external file.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report Header");

        // Bookmark opening.
        builder.Writeln("<<bookmark [\"MyBookmark\"]>>");
        // Insert external document using the supported <<doc>> tag.
        builder.Writeln($"<<doc [\"{Path.GetFileName(externalPath)}\"]>>");
        // Bookmark closing.
        builder.Writeln("<</bookmark>>");

        builder.Writeln("Report Footer");
        templateDoc.Save(templatePath);

        // 3. Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // 4. Build the report (no data needed for this example).
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, new object(), "model");

        // 5. Save the final document.
        string outputPath = Path.Combine(workDir, "output.docx");
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
