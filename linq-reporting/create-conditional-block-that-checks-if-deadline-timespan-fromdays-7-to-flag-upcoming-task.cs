using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin foreach over Tasks collection.
        builder.Writeln("<<foreach [task in Tasks]>>");
        // Task name.
        builder.Writeln("Task: <<[task.Name]>>");
        // Deadline.
        builder.Writeln("Deadline: <<[task.Deadline]>>");
        // Conditional flag for upcoming tasks (deadline within 7 days).
        builder.Writeln("<<if [task.IsUpcoming]>>");
        builder.Writeln("<<textColor [\"Red\"]>>Upcoming<</textColor>>");
        builder.Writeln("<</if>>");
        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Sample data.
        ReportModel model = new()
        {
            Tasks = new()
            {
                new TaskItem { Name = "Task A", Deadline = DateTime.Now.AddDays(3) },
                new TaskItem { Name = "Task B", Deadline = DateTime.Now.AddDays(10) },
                new TaskItem { Name = "Task C", Deadline = DateTime.Now.AddDays(1) }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Root data model.
public class ReportModel
{
    public List<TaskItem> Tasks { get; set; } = new();
}

// Individual task item.
public class TaskItem
{
    public string Name { get; set; } = string.Empty;
    public DateTime Deadline { get; set; }

    // Flag indicating the task deadline is within the next 7 days.
    public bool IsUpcoming => Deadline < DateTime.Now.AddDays(7);
}
