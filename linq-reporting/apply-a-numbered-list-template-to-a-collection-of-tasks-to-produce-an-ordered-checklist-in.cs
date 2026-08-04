using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class TaskItem
{
    // Description of the task.
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
                new() { Description = "Review project requirements" },
                new() { Description = "Design architecture diagram" },
                new() { Description = "Implement core modules" },
                new() { Description = "Write unit tests" },
                new() { Description = "Perform code review" }
            }
        };

        // Create a new blank document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Apply a numbered list style to the paragraph that will contain the LINQ Reporting tags.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Insert the LINQ Reporting tags.
        // <<restartNum>> ensures numbering starts at 1 for the first item.
        // The foreach loop repeats the paragraph for each task in the collection.
        builder.Writeln("<<restartNum>><<foreach [task in Tasks]>><<[task.Description]>>");
        builder.Writeln("<</foreach>>");

        // End the list formatting for subsequent content (optional).
        builder.ListFormat.RemoveNumbers();

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Checklist.docx");
    }
}
