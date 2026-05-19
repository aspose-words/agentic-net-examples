using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class TaskItem
{
    // Name of the task.
    public string Name { get; set; } = string.Empty;

    // Deadline expressed as a TimeSpan.
    public TimeSpan Deadline { get; set; }
}

public class ReportModel
{
    // Collection of tasks to be displayed in the report.
    public List<TaskItem> Tasks { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the temporary template and the final report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Task Report");
        builder.Writeln();

        // Begin a foreach loop over the Tasks collection.
        builder.Writeln("<<foreach [task in Tasks]>>");

        // Output task name.
        builder.Writeln("Name: <<[task.Name]>>");

        // Output deadline (in days) for readability.
        builder.Writeln("Deadline: <<[task.Deadline.Days]>> days");

        // Conditional block: flag tasks whose deadline is less than 7 days.
        builder.Writeln(
            "<<if [task.Deadline.TotalDays < 7]>>" +
            "<<textColor [\"Red\"]>>(Upcoming)<</textColor>><</if>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before building the report).
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and prepare sample data.
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        var model = new ReportModel
        {
            Tasks = new List<TaskItem>
            {
                new TaskItem { Name = "Prepare presentation", Deadline = TimeSpan.FromDays(3) },
                new TaskItem { Name = "Finalize budget", Deadline = TimeSpan.FromDays(10) },
                new TaskItem { Name = "Team meeting", Deadline = TimeSpan.FromDays(5) }
            }
        };

        // -------------------------------------------------
        // 3. Build the report using the ReportingEngine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model); // No root name needed because tags reference members directly.

        // -------------------------------------------------
        // 4. Save the generated report.
        // -------------------------------------------------
        doc.Save(reportPath);
    }
}
