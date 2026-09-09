using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a template document with a numbered list and LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Create a numbered list (default numbering) and apply it to the following paragraphs.
        List list = template.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;

        // Restart numbering before the foreach block.
        builder.Writeln("<<restartNum>><<foreach [task in Tasks]>>");
        // Each paragraph will be a list item showing the task description.
        builder.Writeln("<<[task.Description]>>");
        // End of the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Step 2: Load the template for reporting.
        Document doc = new Document(templatePath);

        // Step 3: Prepare sample data.
        ReportModel model = new ReportModel
        {
            Tasks = new List<TaskItem>
            {
                new TaskItem { Description = "Buy groceries" },
                new TaskItem { Description = "Call the dentist" },
                new TaskItem { Description = "Finish the report" },
                new TaskItem { Description = "Plan weekend trip" }
            }
        };

        // Step 4: Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(doc, model, "model");

        // Step 5: Save the generated checklist.
        const string outputPath = "Checklist.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated successfully: {outputPath}");
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<TaskItem> Tasks { get; set; } = new();
}

// Individual task item.
public class TaskItem
{
    public string Description { get; set; } = "";
}
