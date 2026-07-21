using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class TaskItem
{
    public string Name { get; set; } = "";
    public TimeSpan Deadline { get; set; }
}

public class ReportModel
{
    public List<TaskItem> Tasks { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var model = new ReportModel
        {
            Tasks = new List<TaskItem>
            {
                new TaskItem { Name = "Design document", Deadline = TimeSpan.FromDays(3) },
                new TaskItem { Name = "Code implementation", Deadline = TimeSpan.FromDays(10) },
                new TaskItem { Name = "Testing", Deadline = TimeSpan.FromDays(5) }
            }
        };

        // Create a template document.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Task Report");
        builder.Writeln("--------------------");

        // LINQ Reporting tags.
        builder.Writeln("<<foreach [task in Tasks]>>");
        builder.Writeln("Task: <<[task.Name]>>");
        builder.Writeln("Deadline (days): <<[task.Deadline.Days]>>");
        // Flag tasks whose deadline is less than 7 days.
        builder.Writeln("<<if [task.Deadline.TotalDays < 7]>>Upcoming<</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        reportDoc.Save("TaskReport.docx");
    }
}
