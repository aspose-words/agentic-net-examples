using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure code page support (required by Aspose.Words for some encodings).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 1. Create sample XML data representing tasks.
        const string xmlFileName = "tasks.xml";
        string xmlContent =
            @"<?xml version=""1.0"" encoding=""UTF-8""?>
              <tasks>
                  <task>
                      <title>Buy groceries</title>
                      <description>Milk, Bread, Eggs</description>
                  </task>
                  <task>
                      <title>Call Alice</title>
                      <description>Discuss project timeline</description>
                  </task>
                  <task>
                      <title>Write report</title>
                      <description>Annual financial summary</description>
                  </task>
              </tasks>";
        File.WriteAllText(xmlFileName, xmlContent);

        // 2. Build the template document programmatically.
        const string templateFileName = "TaskTemplate.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title of the report.
        builder.Writeln("Task List:");

        // Start a bulleted list.
        builder.ListFormat.ApplyBulletDefault();

        // LINQ Reporting foreach tag iterating over the XML collection "tasks".
        builder.Writeln("<<foreach [task in tasks]>>");
        // Each bullet will contain the task title and description.
        builder.Writeln("<<[task.title]>>: <<[task.description]>>");
        builder.Writeln("<</foreach>>");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the template to disk.
        templateDoc.Save(templateFileName);

        // 3. Load the template for report generation.
        var reportDoc = new Document(templateFileName);

        // 4. Create an XmlDataSource from the XML file.
        var xmlDataSource = new XmlDataSource(xmlFileName);

        // 5. Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        // The data source name must match the name used in the template ("tasks").
        engine.BuildReport(reportDoc, xmlDataSource, "tasks");

        // 6. Save the generated report.
        const string outputFileName = "TaskReport.docx";
        reportDoc.Save(outputFileName);
    }
}
