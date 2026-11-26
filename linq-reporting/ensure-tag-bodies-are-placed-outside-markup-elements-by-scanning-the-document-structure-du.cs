using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model.
        var model = new ReportModel
        {
            Name = "John Doe"
        };

        // Create a template document with a tag inside a paragraph that also contains other text.
        string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Preprocess the document to ensure tag bodies are placed outside markup elements.
        PreprocessDocument(doc);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }

    // Creates a simple template containing a tag inside a paragraph with surrounding text.
    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph with tag embedded in text.
        builder.Writeln("Hello <<[model.Name]>> world!");

        // Save the template.
        doc.Save(path);
    }

    // Scans the document and moves tag bodies to separate runs if they share a run with other text.
    private static void PreprocessDocument(Document doc)
    {
        // Regular expression to find LINQ Reporting tags like <<[...]>>
        Regex tagRegex = new Regex(@"<<\[[^\]]+\]>>", RegexOptions.Compiled);

        // Collect runs to process to avoid modifying the collection while iterating.
        List<Run> runs = new List<Run>();
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            runs.Add(run);
        }

        foreach (Run run in runs)
        {
            string text = run.Text;
            Match match = tagRegex.Match(text);
            if (match.Success && match.Length != text.Length)
            {
                string before = text.Substring(0, match.Index);
                string tag = match.Value;
                string after = text.Substring(match.Index + match.Length);

                // Insert runs in the order: before, tag, after.
                if (!string.IsNullOrEmpty(before))
                {
                    Run beforeRun = (Run)run.Clone(true);
                    beforeRun.Text = before;
                    run.ParentNode.InsertBefore(beforeRun, run);
                }

                Run tagRun = (Run)run.Clone(true);
                tagRun.Text = tag;
                run.ParentNode.InsertBefore(tagRun, run);

                if (!string.IsNullOrEmpty(after))
                {
                    Run afterRun = (Run)run.Clone(true);
                    afterRun.Text = after;
                    run.ParentNode.InsertBefore(afterRun, run);
                }

                // Remove the original mixed run.
                run.Remove();
            }
        }
    }

    // Sample data model used by the report.
    public class ReportModel
    {
        public string Name { get; set; } = string.Empty;
    }
}
