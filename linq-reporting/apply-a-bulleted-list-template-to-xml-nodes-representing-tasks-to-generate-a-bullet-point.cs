using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for XML handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample XML data.
        string xmlPath = "tasks.xml";
        File.WriteAllText(xmlPath, @"<?xml version=""1.0"" encoding=""utf-8""?>
<Tasks>
    <Task><Title>Review project proposal</Title></Task>
    <Task><Title>Design architecture diagram</Title></Task>
    <Task><Title>Implement core module</Title></Task>
    <Task><Title>Write unit tests</Title></Task>
    <Task><Title>Prepare release notes</Title></Task>
</Tasks>");

        // Create the template document programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a paragraph that will become a bulleted list via LINQ Reporting.
        builder.Writeln("<<foreach [task in Tasks]>>• <<[task.Title]>> <</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Create XML data source.
        XmlDataSource xmlData = new XmlDataSource(xmlPath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(reportDoc, xmlData, "Tasks");

        // Save the generated report.
        string outputPath = "output.docx";
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }
}
