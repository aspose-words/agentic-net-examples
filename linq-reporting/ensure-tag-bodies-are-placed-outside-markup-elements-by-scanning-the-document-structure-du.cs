using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingTagPreprocess
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "World";
    }

    public class Program
    {
        // Entry point.
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // The tag is placed inside a sentence – this is the situation we want to fix.
            builder.Writeln("Hello <<[model.Name]>>! This is a LINQ Reporting example.");

            // Save the template so it can be re‑loaded for preprocessing.
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back and preprocess it.
            //    The goal is to ensure that any tag body (e.g., <<[model.Name]>>)
            //    is not embedded inside other markup elements such as runs.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);
            MoveTagsToSeparateParagraphs(doc);

            // -----------------------------------------------------------------
            // 3. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            var model = new ReportModel(); // model.Name = "World" by default

            var engine = new ReportingEngine
            {
                // Use the property setter as required by the rules.
                Options = ReportBuildOptions.None
            };

            // The root name in the template is "model" because the tag uses <<[model.Name]>>.
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(reportPath);
        }

        // Scans the document and moves any tag found inside a Run to its own paragraph.
        private static void MoveTagsToSeparateParagraphs(Document document)
        {
            // Create a snapshot of all runs to avoid modifying the collection while iterating.
            var runs = document.GetChildNodes(NodeType.Run, true)
                               .Cast<Run>()
                               .ToList();

            foreach (Run run in runs)
            {
                string text = run.Text;
                int startIdx = text.IndexOf("<<[", StringComparison.Ordinal);
                int endIdx = text.IndexOf("]>>", StringComparison.Ordinal);

                // If a tag is present inside this run, extract it.
                if (startIdx >= 0 && endIdx > startIdx)
                {
                    // Preserve any text before the tag.
                    string beforeTag = text.Substring(0, startIdx);
                    // Preserve any text after the tag.
                    string afterTag = text.Substring(endIdx + 3);

                    // Replace the original run's text with the preceding text (if any).
                    run.Text = beforeTag;

                    // Create a new paragraph that will contain only the tag.
                    Paragraph currentParagraph = (Paragraph)run.GetAncestor(NodeType.Paragraph);
                    Paragraph tagParagraph = (Paragraph)currentParagraph.Clone(true);
                    tagParagraph.Runs.Clear();

                    // Create a new run that contains only the tag.
                    Run tagRun = new Run(document, text.Substring(startIdx, endIdx - startIdx + 3));
                    tagParagraph.Runs.Add(tagRun);

                    // Insert the tag paragraph immediately after the current paragraph.
                    CompositeNode parent = (CompositeNode)currentParagraph.ParentNode;
                    parent.InsertAfter(tagParagraph, currentParagraph);

                    // If there is trailing text after the tag, insert it as a new run in the original paragraph.
                    if (!string.IsNullOrEmpty(afterTag))
                    {
                        Run afterRun = new Run(document, afterTag);
                        currentParagraph.Runs.Add(afterRun);
                    }
                }
            }
        }
    }
}
