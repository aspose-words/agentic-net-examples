using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace LinqReportingTagPreprocess
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for certain data sources.
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Sample data model.
            Order order = new Order { CustomerName = "John Doe" };

            // Create a template document with a LINQ Reporting tag inside a run.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("Dear <<[order.CustomerName]>>,");
            builder.Writeln("Thank you for your purchase.");

            // Save and reload to emulate a real‑world scenario.
            const string templatePath = "Template.docx";
            template.Save(templatePath);
            Document doc = new Document(templatePath);

            // Preprocess: ensure each tag resides in its own Run node.
            PreprocessDocument(doc);

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, order, "order");

            // Save the final document.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }

        // Moves any LINQ Reporting tags so that they are isolated in separate Run nodes.
        private static void PreprocessDocument(Document doc)
        {
            foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                // Get a snapshot of the runs in the paragraph.
                Run[] runs = paragraph.GetChildNodes(NodeType.Run, true)
                                      .OfType<Run>()
                                      .ToArray();

                foreach (Run run in runs)
                {
                    string text = run.Text;
                    int startIdx = text.IndexOf("<<[", StringComparison.Ordinal);
                    int endIdx = text.IndexOf("]>>", StringComparison.Ordinal);

                    // If a tag is found, split the run into before‑tag, tag, and after‑tag parts.
                    if (startIdx >= 0 && endIdx > startIdx)
                    {
                        string beforeTag = text.Substring(0, startIdx);
                        string tag = text.Substring(startIdx, endIdx - startIdx + 3);
                        string afterTag = text.Substring(endIdx + 3);

                        // Preserve the original run for insertion reference.
                        Run referenceRun = run;

                        // Set the text of the original run to the preceding text (may be empty).
                        referenceRun.Text = beforeTag ?? string.Empty;

                        // Insert a new run that contains only the tag.
                        Run tagRun = new Run(doc, tag);
                        paragraph.InsertAfter(tagRun, referenceRun);

                        // If there is trailing text, insert another run after the tag run.
                        if (!string.IsNullOrEmpty(afterTag))
                        {
                            Run afterRun = new Run(doc, afterTag);
                            paragraph.InsertAfter(afterRun, tagRun);
                        }

                        // If the original run ended up empty, remove it to avoid stray empty runs.
                        if (string.IsNullOrEmpty(referenceRun.Text))
                        {
                            referenceRun.Remove();
                        }
                    }
                }
            }
        }
    }

    // Public data model used by the template.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
    }
}
