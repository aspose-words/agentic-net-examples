using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Model representing a single task.
    public class TaskItem
    {
        public string Description { get; set; } = string.Empty;
    }

    // Wrapper model that will be passed to the ReportingEngine.
    public class ReportModel
    {
        public List<TaskItem> Tasks { get; set; } = new();
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            // Paths for the temporary template and the final report.
            string templatePath = "ChecklistTemplate.docx";
            string outputPath = "ChecklistReport.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Create a numbered list style.
            List numberedList = templateDoc.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = numberedList;

            // Insert a single paragraph that contains the LINQ Reporting tags.
            // <<restartNum>> ensures numbering starts at 1.
            // The foreach loop repeats the paragraph for each task in the collection.
            builder.Writeln("<<restartNum>><<foreach [task in Tasks]>> <<[task.Description]>> <</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Sample data: a collection of tasks.
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

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated checklist.
            reportDoc.Save(outputPath);
        }
    }
}
