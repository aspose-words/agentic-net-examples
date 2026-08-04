using System;
using System.Collections.Generic;
using System.IO;
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
        // Register code page provider for Aspose.Words (required for some encodings)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Step 1: Create the template document programmatically
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin a foreach loop over the Tasks collection
        builder.Writeln("<<foreach [task in Tasks]>>");
        // Output the task name
        builder.Writeln("Task: <<[task.Name]>>");
        // Conditional block: flag tasks with a deadline less than 7 days
        builder.Writeln("<<if [task.IsUpcoming]>>");
        builder.Writeln(" - Upcoming (deadline in <<[task.Deadline]>>)");
        builder.Writeln("<</if>>");
        // End the foreach loop
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        doc.Save(templatePath);

        // Step 2: Load the template document for reporting
        var templateDoc = new Document(templatePath);

        // Step 3: Prepare sample data
        var model = new ReportModel
        {
            Tasks = new List<TaskItem>
            {
                new TaskItem { Name = "Prepare presentation", Deadline = TimeSpan.FromDays(3) },
                new TaskItem { Name = "Submit report", Deadline = TimeSpan.FromDays(10) },
                new TaskItem { Name = "Team meeting", Deadline = TimeSpan.FromDays(5) }
            }
        };

        // Step 4: Build the report using the LINQ Reporting engine
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Step 5: Save the generated report
        var outputPath = "Report.docx";
        templateDoc.Save(outputPath);
    }
}
