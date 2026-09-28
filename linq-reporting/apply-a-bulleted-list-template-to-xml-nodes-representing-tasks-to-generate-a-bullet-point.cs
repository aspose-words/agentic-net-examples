using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample XML data representing tasks.
        const string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Tasks>
    <Task>
        <Title>Prepare project proposal</Title>
    </Task>
    <Task>
        <Title>Review design specifications</Title>
    </Task>
    <Task>
        <Title>Implement core modules</Title>
    </Task>
    <Task>
        <Title>Test and validate features</Title>
    </Task>
</Tasks>";
        const string xmlPath = "tasks.xml";
        File.WriteAllText(xmlPath, xmlContent);

        // Create the LINQ Reporting template programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Optional title.
        builder.Writeln("Task List:");
        builder.Writeln();

        // Begin foreach loop over XML nodes.
        builder.Writeln("<<foreach [task in Tasks]>>");

        // Apply bullet list formatting for each task.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("<<[task.Title]>>");
        // Clear list formatting so it does not affect following paragraphs.
        builder.ListFormat.RemoveNumbers();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template for report generation.
        Document report = new Document(templatePath);

        // Load XML data source.
        XmlDataSource xmlData = new XmlDataSource(xmlPath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, xmlData, "Tasks");

        // Save the generated report.
        const string outputPath = "output.docx";
        report.Save(outputPath);
    }
}
