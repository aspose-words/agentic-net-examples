using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model for the report.
    public class TaskItem
    {
        public int Index { get; set; }
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
            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Apply a numbered list style to the paragraph that will contain the items.
            List numberedList = template.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = numberedList;

            // Insert the LINQ Reporting tags.
            // <<restartNum>> resets the numbering at the start of the list.
            // <<foreach [task in Tasks]>> iterates over the collection.
            // Each iteration writes the task description on a new paragraph.
            builder.Writeln("<<restartNum>><<foreach [task in Tasks]>><<[task.Description]>>\r<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document report = new Document(templatePath);

            // Prepare sample data.
            ReportModel model = new ReportModel
            {
                Tasks = new List<TaskItem>
                {
                    new TaskItem { Index = 1, Description = "Review project requirements" },
                    new TaskItem { Index = 2, Description = "Design architecture" },
                    new TaskItem { Index = 3, Description = "Implement core modules" },
                    new TaskItem { Index = 4, Description = "Write unit tests" },
                    new TaskItem { Index = 5, Description = "Perform code review" }
                }
            };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options
            engine.BuildReport(report, model, "model");

            // Save the generated checklist.
            const string outputPath = "Checklist.docx";
            report.Save(outputPath);
        }
    }
}
