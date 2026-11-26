using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalExample
{
    // Root data model
    public class ReportModel
    {
        public List<TaskItem> Tasks { get; set; } = new();
    }

    // Individual task item
    public class TaskItem
    {
        public string Name { get; set; } = "";
        public DateTime Deadline { get; set; }

        // Calculated property used in the conditional tag
        public TimeSpan TimeUntilDeadline => Deadline - DateTime.Now;

        // Helper property for the conditional block (deadline within 7 days)
        public bool IsUpcoming => TimeUntilDeadline.TotalDays < 7;
    }

    class Program
    {
        static void Main()
        {
            // Register code page provider (required by Aspose.Words for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare template file
            const string templatePath = "template.docx";
            CreateTemplate(templatePath);

            // Load the template
            Document doc = new Document(templatePath);

            // Prepare sample data
            ReportModel model = new()
            {
                Tasks = new()
                {
                    new TaskItem { Name = "Task A", Deadline = DateTime.Now.AddDays(3) },
                    new TaskItem { Name = "Task B", Deadline = DateTime.Now.AddDays(10) },
                    new TaskItem { Name = "Task C", Deadline = DateTime.Now.AddDays(1) }
                }
            };

            // Build the report
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the result
            const string outputPath = "report.docx";
            doc.Save(outputPath);
        }

        // Creates a simple Word template with LINQ Reporting tags
        private static void CreateTemplate(string path)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Task Report");
            builder.Writeln("--------------------");

            // Begin foreach over Tasks collection
            builder.Writeln("<<foreach [task in Tasks]>>");

            // Task name
            builder.Writeln("Task: <<[task.Name]>>");

            // Deadline display
            builder.Writeln("Deadline: <<[task.Deadline]>>");

            // Conditional block: flag if deadline is within 7 days
            builder.Writeln("<<if [task.IsUpcoming]>>");
            builder.Writeln("** Upcoming! **");
            builder.Writeln("<</if>>");

            // End foreach
            builder.Writeln("<</foreach>>");

            // Save the template
            doc.Save(path);
        }
    }
}
