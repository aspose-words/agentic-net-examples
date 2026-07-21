using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Sample XML representing a list of tasks.
        const string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Tasks>
    <Task><Name>Buy groceries</Name></Task>
    <Task><Name>Call the bank</Name></Task>
    <Task><Name>Finish the report</Name></Task>
    <Task><Name>Schedule meeting</Name></Task>
</Tasks>";

        // Write XML to a memory stream.
        using var xmlStream = new MemoryStream();
        using (var writer = new StreamWriter(xmlStream, System.Text.Encoding.UTF8, leaveOpen: true))
        {
            writer.Write(xmlContent);
        }
        xmlStream.Position = 0; // Reset for reading.

        // -----------------------------------------------------------------
        // Create the template document.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Apply a bulleted list style to the document.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList; // Subsequent paragraphs become list items.

        // Insert LINQ Reporting tags.
        // The XML root <Tasks> contains a collection of <Task> elements.
        builder.Writeln("<<foreach [task in Tasks]>>");
        builder.Writeln("<<[task.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, shown for clarity).
        const string templatePath = "Template.docx";
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Reset the XML stream before creating the data source.
        xmlStream.Position = 0;
        var xmlDataSource = new XmlDataSource(xmlStream);

        // Build the report. Provide a data source name that matches the root element.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, xmlDataSource, "Tasks");

        // Save the generated report.
        const string outputPath = "TaskListReport.docx";
        reportDoc.Save(outputPath);
    }
}
