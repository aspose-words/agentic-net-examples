using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        ReportModel model = new()
        {
            Tasks = new List<TaskItem>
            {
                new() { Name = "Prepare presentation", Deadline = TimeSpan.FromDays(3) },
                new() { Name = "Submit report", Deadline = TimeSpan.FromDays(10) },
                new() { Name = "Team meeting", Deadline = TimeSpan.FromDays(5) }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new();
        DocumentBuilder builder = new(template);

        builder.Writeln("Task Report");
        builder.Writeln("<<foreach [task in Tasks]>>");
        builder.Writeln("Name: <<[task.Name]>>");
        // Use TotalDays to avoid static method calls that the engine cannot resolve.
        builder.Writeln("<<if [task.Deadline.TotalDays < 7]>>");
        builder.Writeln(" - Upcoming (deadline in <<[task.Deadline]>>)");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}

// Wrapper class for the root data source.
public class ReportModel
{
    public List<TaskItem> Tasks { get; set; } = new();
}

// Individual task item.
public class TaskItem
{
    public string Name { get; set; } = string.Empty;
    public TimeSpan Deadline { get; set; }
}
