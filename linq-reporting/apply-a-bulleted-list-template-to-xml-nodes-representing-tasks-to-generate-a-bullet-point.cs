using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        const string templatePath = "TaskTemplate.docx";
        const string reportPath = "TaskReport.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Create a bulleted list that will be used for each task.
        List bulletList = templateDoc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // Insert LINQ Reporting tags.
        // The XML data source will expose a collection named "tasks".
        builder.Writeln("<<foreach [task in tasks]>>");
        // Each iteration writes the task title as a list item.
        builder.Writeln("<<[task.Title]>>");
        builder.Writeln("<</foreach>>");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template for report generation.
        // -------------------------------------------------
        var reportDoc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare sample XML data representing tasks.
        // -------------------------------------------------
        const string xmlContent = @"
<tasks>
    <task>
        <Title>Buy groceries</Title>
    </task>
    <task>
        <Title>Call Alice</Title>
    </task>
    <task>
        <Title>Finish project report</Title>
    </task>
</tasks>";
        using var xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlContent));

        // Create an XmlDataSource from the XML stream.
        var xmlDataSource = new XmlDataSource(xmlStream);

        // -------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        // -------------------------------------------------
        var engine = new ReportingEngine();
        // The root object name in the template is "tasks".
        engine.BuildReport(reportDoc, xmlDataSource, "tasks");

        // -------------------------------------------------
        // 5. Save the generated report.
        // -------------------------------------------------
        reportDoc.Save(reportPath);
    }
}
