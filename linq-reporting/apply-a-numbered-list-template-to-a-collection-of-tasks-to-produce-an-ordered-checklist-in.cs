using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class TaskItem
{
    public string Description { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<TaskItem> Tasks { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Tasks = new List<TaskItem>
            {
                new TaskItem { Description = "Review project requirements" },
                new TaskItem { Description = "Design architecture diagram" },
                new TaskItem { Description = "Implement core modules" },
                new TaskItem { Description = "Write unit tests" },
                new TaskItem { Description = "Perform code review" }
            }
        };

        // Create template document with LINQ Reporting tags.
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a numbered paragraph that will be repeated for each task.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("<<restartNum>><<foreach [task in Tasks]>><<[task.Description]>> <</foreach>>");
        builder.ListFormat.RemoveNumbers();

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        const string outputPath = "Checklist.docx";
        doc.Save(outputPath);
    }
}
