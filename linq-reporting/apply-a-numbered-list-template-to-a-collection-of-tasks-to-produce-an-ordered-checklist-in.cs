using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // 1. Create a template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Apply a numbered list style to the following paragraphs.
        List numberedList = template.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;

        // Insert LINQ Reporting tags.
        // <<restartNum>> placed before <<foreach>> restarts numbering for the list.
        builder.Writeln("<<restartNum>><<foreach [task in Tasks]>>");
        builder.Writeln("<<[task.Title]>>");
        builder.Writeln("<</foreach>>");

        // 2. Prepare sample data.
        ReportModel model = new ReportModel
        {
            Tasks = new List<TaskItem>
            {
                new TaskItem { Title = "Buy groceries" },
                new TaskItem { Title = "Call Alice" },
                new TaskItem { Title = "Finish report" },
                new TaskItem { Title = "Schedule meeting" }
            }
        };

        // 3. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        bool success = engine.BuildReport(template, model, "model");

        // 4. Save the generated document.
        template.Save("ChecklistReport.docx");
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<TaskItem> Tasks { get; set; } = new();
}

// Simple task item model.
public class TaskItem
{
    public string Title { get; set; } = string.Empty;
}
