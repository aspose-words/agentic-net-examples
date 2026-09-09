using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class TaskItem
{
    public string Name { get; set; } = string.Empty;
    public TimeSpan Deadline { get; set; }

    // Helper property used in the template to flag upcoming tasks.
    public bool IsUpcoming => Deadline < TimeSpan.FromDays(7);
}

public class ReportModel
{
    public List<TaskItem> Tasks { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // ---------- Create template ----------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin foreach over Tasks
        builder.Writeln("<<foreach [task in Tasks]>>");
        // Write task name
        builder.Writeln("Task: <<[task.Name]>>");
        // Conditional block: flag tasks with deadline less than 7 days
        builder.Writeln("<<if [task.IsUpcoming]>>");
        builder.Writeln(" - Upcoming!");
        builder.Writeln("<</if>>");
        // End foreach
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // ---------- Load template ----------
        var doc = new Document(templatePath);

        // ---------- Prepare data ----------
        var model = new ReportModel
        {
            Tasks = new List<TaskItem>
            {
                new TaskItem { Name = "Prepare report", Deadline = TimeSpan.FromDays(5) },
                new TaskItem { Name = "Finalize budget", Deadline = TimeSpan.FromDays(10) },
                new TaskItem { Name = "Team meeting", Deadline = TimeSpan.FromDays(2) }
            }
        };

        // ---------- Build report ----------
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // ---------- Save output ----------
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
