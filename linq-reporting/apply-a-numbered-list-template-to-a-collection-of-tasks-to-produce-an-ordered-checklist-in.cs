using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Model representing a single task.
    public class TaskItem
    {
        public int Index { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    // Wrapper model that will be passed to the ReportingEngine.
    public class ReportModel
    {
        public List<TaskItem> Tasks { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Create a numbered list style that will be applied to the paragraph containing the tags.
            List numberedList = template.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = numberedList;

            // Insert a heading.
            builder.Writeln("Project Checklist:");

            // Insert a numbered paragraph that will be repeated for each task.
            // The <<restartNum>> tag ensures numbering starts from 1 for this list.
            // The foreach tag iterates over the Tasks collection.
            builder.Writeln("<<restartNum>><<foreach [task in Tasks]>><<[task.Index]>>. <<[task.Description]>> <</foreach>>");

            // Save the template to disk (required before building the report).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Prepare sample data.
            var model = new ReportModel
            {
                Tasks = new List<TaskItem>
                {
                    new TaskItem { Index = 1, Description = "Gather requirements" },
                    new TaskItem { Index = 2, Description = "Design architecture" },
                    new TaskItem { Index = 3, Description = "Implement features" },
                    new TaskItem { Index = 4, Description = "Write unit tests" },
                    new TaskItem { Index = 5, Description = "Perform code review" }
                }
            };

            // 3. Load the template and build the report.
            Document report = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            // No special options are needed for this simple scenario.
            engine.BuildReport(report, model, "model");

            // 4. Save the generated report.
            const string outputPath = "ChecklistReport.docx";
            report.Save(outputPath);

            // Inform the user (optional, no interactive input required).
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
