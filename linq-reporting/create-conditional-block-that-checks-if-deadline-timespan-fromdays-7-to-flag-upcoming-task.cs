using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class TaskInfo
{
    // Name of the task.
    public string Name { get; set; } = string.Empty;

    // Deadline expressed as a TimeSpan (e.g., days remaining).
    public TimeSpan Deadline { get; set; }

    // Helper property used by the template to avoid '<' inside the condition.
    public bool IsUpcoming => Deadline.TotalDays < 7;
}

public class ReportModel
{
    // Collection of tasks to be processed by the LINQ Reporting engine.
    public List<TaskInfo> Tasks { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Tasks Report");
        builder.Writeln(); // Empty line for readability.

        // Begin a foreach loop over the Tasks collection of the root model object.
        builder.Writeln("<<foreach [task in model.Tasks]>>");

        // Output task name.
        builder.Writeln("Task: <<[task.Name]>>");

        // Output the raw deadline value.
        builder.Writeln("Deadline: <<[task.Deadline]>>");

        // Conditional block: flag tasks whose deadline is less than 7 days.
        // The condition uses the helper property IsUpcoming to avoid '<' inside the tag.
        builder.Writeln("<<if [task.IsUpcoming]>>Upcoming!<</if>>");

        // End of foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for report generation.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare sample data.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Tasks = new List<TaskInfo>
            {
                new TaskInfo { Name = "Prepare presentation", Deadline = TimeSpan.FromDays(3) },
                new TaskInfo { Name = "Finalize budget", Deadline = TimeSpan.FromDays(10) },
                new TaskInfo { Name = "Team meeting", Deadline = TimeSpan.FromDays(5) }
            }
        };

        // -----------------------------------------------------------------
        // 4. Build the report using Aspose.Words LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // Allow the template to use static members of TimeSpan if needed.
        engine.KnownTypes.Add(typeof(TimeSpan));

        // Build the report. The root object name in the template is "model".
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);

        // Indicate completion (no interactive prompts).
        Console.WriteLine($"Report generated at: {reportPath}");
    }
}
