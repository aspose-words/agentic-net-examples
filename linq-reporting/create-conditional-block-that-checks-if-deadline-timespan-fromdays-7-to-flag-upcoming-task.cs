using System;
using System.Collections.Generic;
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
        // Create the template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("<<foreach [t in Tasks]>>");
        builder.Writeln("Task: <<[t.Name]>>");
        builder.Writeln("Deadline: <<[t.Deadline]>>");
        builder.Writeln("<<if [t.IsUpcoming]>>Upcoming!<</if>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Tasks = new()
            {
                new TaskItem { Name = "Prepare presentation", Deadline = TimeSpan.FromDays(3) },
                new TaskItem { Name = "Submit report", Deadline = TimeSpan.FromDays(10) },
                new TaskItem { Name = "Team meeting", Deadline = TimeSpan.FromDays(5) }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
