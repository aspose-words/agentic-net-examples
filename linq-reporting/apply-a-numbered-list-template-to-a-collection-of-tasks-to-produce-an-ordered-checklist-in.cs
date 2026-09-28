using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class TaskItem
{
    public string Description { get; set; } = string.Empty;
    public bool IsCompleted { get; set; }
}

public class ReportModel
{
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
                new TaskItem { Description = "Review project requirements", IsCompleted = false },
                new TaskItem { Description = "Design architecture diagram", IsCompleted = true },
                new TaskItem { Description = "Implement core modules", IsCompleted = false },
                new TaskItem { Description = "Write unit tests", IsCompleted = false },
                new TaskItem { Description = "Perform code review", IsCompleted = true }
            }
        };

        // Paths for template and output.
        string templatePath = "ChecklistTemplate.docx";
        string outputPath = "ChecklistResult.docx";

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.Writeln("Project Checklist");
        builder.ParagraphFormat.ClearFormatting();

        // Start a numbered list for the tasks.
        builder.ListFormat.ApplyNumberDefault();

        // Insert LINQ Reporting tags.
        // <<restartNum>> placed in the same numbered paragraph before the foreach.
        builder.Writeln("<<restartNum>><<foreach [task in Tasks]>><<[task.Description]>> <</foreach>>");

        // End the numbered list.
        builder.ListFormat.RemoveNumbers();

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save(outputPath);
    }
}
