using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    // Simple data model used by the LINQ Reporting template.
    public class Model
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
    }

    public static void Main()
    {
        // Step 1: Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Write some text, then start a bold run that incorrectly contains a tag body.
        builder.Writeln("Report for:");
        builder.Font.Bold = true;
        builder.Write("<<[model.Name]>>"); // Tag inside a bold run – this is the situation we want to fix.
        builder.Font.Bold = false;
        builder.Writeln(" generated on " + DateTime.Now.ToShortDateString());

        // Save the template (optional, just to visualize the original state).
        template.Save("Template.docx");

        // Step 2: Preprocess the document to ensure tag bodies are placed outside markup elements.
        PreprocessTags(template);

        // Save the preprocessed template to verify the tags have been moved.
        template.Save("PreprocessedTemplate.docx");

        // Step 3: Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this simple example.
        Model data = new Model();
        engine.BuildReport(template, data, "model");

        // Step 4: Save the final report.
        template.Save("Report.docx");
    }

    // Scans the document for Run nodes that contain LINQ Reporting tags and moves those tags
    // into separate runs without the original formatting (e.g., bold, italic).
    private static void PreprocessTags(Document doc)
    {
        // Collect runs that need processing to avoid modifying the collection while iterating.
        List<Run> runsToProcess = new List<Run>();
        NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
        foreach (Run run in runs)
        {
            if (run.Text.Contains("<<"))
                runsToProcess.Add(run);
        }

        foreach (Run run in runsToProcess)
        {
            string text = run.Text;
            int startIdx = text.IndexOf("<<", StringComparison.Ordinal);
            int endIdx = text.IndexOf(">>", startIdx, StringComparison.Ordinal);
            if (startIdx < 0 || endIdx < 0)
                continue; // No well‑formed tag found.

            // Extract parts surrounding the tag.
            string before = text.Substring(0, startIdx);
            string tag = text.Substring(startIdx, endIdx - startIdx + 2); // include ">>"
            string after = text.Substring(endIdx + 2);

            Paragraph paragraph = (Paragraph)run.ParentNode;
            // Insert runs in the correct order: before text, tag, after text.
            // Insert before the original run so we can later remove it.
            if (!string.IsNullOrEmpty(before))
            {
                Run beforeRun = (Run)run.Clone(true);
                beforeRun.Text = before;
                paragraph.InsertBefore(beforeRun, run);
            }

            // Tag run – clear formatting to ensure it is outside markup.
            Run tagRun = (Run)run.Clone(true);
            tagRun.Text = tag;
            // Reset formatting (remove bold, italic, etc.).
            tagRun.Font.Bold = false;
            tagRun.Font.Italic = false;
            tagRun.Font.Underline = Underline.None;
            paragraph.InsertBefore(tagRun, run);

            if (!string.IsNullOrEmpty(after))
            {
                Run afterRun = (Run)run.Clone(true);
                afterRun.Text = after;
                paragraph.InsertBefore(afterRun, run);
            }

            // Remove the original run that contained the mixed content.
            run.Remove();
        }
    }
}
