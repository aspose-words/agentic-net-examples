using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class TaskItem
{
    // Description of the task; initialized to avoid nullable warnings.
    public string Description { get; set; } = string.Empty;
}

public class ReportModel
{
    // Collection of tasks to be listed.
    public List<TaskItem> Tasks { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Tasks = new List<TaskItem>
            {
                new TaskItem { Description = "Buy groceries" },
                new TaskItem { Description = "Call the dentist" },
                new TaskItem { Description = "Finish the report" },
                new TaskItem { Description = "Plan weekend trip" }
            }
        };

        // Create the template document programmatically.
        const string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Task Checklist:");
        builder.Writeln(); // Blank line for spacing.

        // Apply a numbered list style to the upcoming paragraphs.
        builder.ListFormat.ApplyNumberDefault();

        // LINQ Reporting tags:
        // Open the foreach loop.
        builder.Writeln("<<foreach [task in Tasks]>>");
        // Each task description is a numbered paragraph; restart numbering for each item.
        builder.Writeln("<<restartNum>><<[task.Description]>>");
        // Close the foreach loop.
        builder.Writeln("<</foreach>>");

        // Remove list formatting after the loop.
        builder.ListFormat.RemoveNumbers();

        // Save the template to disk.
        doc.Save(templatePath);

        // Load the template back (ensures the document is fully prepared before building the report).
        var templateDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None // No special options needed.
        };
        bool success = engine.BuildReport(templateDoc, model, "model");

        // Save the generated report if the build succeeded.
        if (success)
        {
            const string outputPath = "Checklist.docx";
            templateDoc.Save(outputPath);
        }
    }
}
