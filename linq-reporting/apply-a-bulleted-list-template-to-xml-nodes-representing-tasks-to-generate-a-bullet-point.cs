using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class TaskItem
{
    public string Title { get; set; } = string.Empty;
}

public class TaskModel
{
    public List<TaskItem> Task { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Prepare sample data.
        // -----------------------------------------------------------------
        var model = new TaskModel
        {
            Task =
            {
                new TaskItem { Title = "Review project requirements" },
                new TaskItem { Title = "Design database schema" },
                new TaskItem { Title = "Implement authentication module" },
                new TaskItem { Title = "Write unit tests" }
            }
        };

        // -----------------------------------------------------------------
        // 2. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title paragraph.
        builder.Writeln("Task List:");
        builder.Writeln();

        // Apply bullet formatting to the paragraph that will contain the foreach block.
        builder.ListFormat.ApplyBulletDefault();

        // Begin foreach loop over Task items.
        builder.Writeln("<<foreach [t in Task]>>");
        // Write the task title.
        builder.Writeln("<<[t.Title]>>");
        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "TaskTemplate.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and generate the report using the data model.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report; "model" is the name of the root data source.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "TaskReport.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
