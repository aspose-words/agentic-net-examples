using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple XML file with a list of tasks.
        const string xmlFileName = "Tasks.xml";
        File.WriteAllText(xmlFileName,
@"<Tasks>
    <Task><Name>Buy groceries</Name></Task>
    <Task><Name>Call Alice</Name></Task>
    <Task><Name>Finish report</Name></Task>
</Tasks>");

        // Create a blank document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Apply a bulleted list style to subsequent paragraphs.
        List bulletList = template.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // Insert LINQ Reporting tags.
        // The foreach block iterates over the XML nodes and writes each task name.
        builder.Writeln("<<foreach [task in Tasks]>>");
        builder.Writeln("<<[task.Name]>>");
        builder.Writeln("<</foreach>>");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Load the XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlFileName);

        // Build the report by merging the template with the XML data.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSource, "Tasks");

        // Save the generated report.
        const string outputFileName = "TaskReport.docx";
        template.Save(outputFileName);
    }
}
