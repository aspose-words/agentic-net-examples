using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingTagPreprocess
{
    // Simple data model used by the template.
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public List<Item> Items { get; set; } = new()
        {
            new Item { Index = 1, Name = "Apple" },
            new Item { Index = 2, Name = "Banana" }
        };
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Paragraph with a tag inside a run (simulating a malformed placement).
            builder.Writeln("Order for <<[order.CustomerName]>>:");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln(" - Item <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "TagTemplate.docx";
            template.Save(templatePath);

            // 2. Load the template back for preprocessing.
            Document doc = new Document(templatePath);

            // 3. Preprocess: ensure that tag bodies are placed outside markup elements.
            //    For each paragraph, if a run contains a tag (<<...>>), move the whole tag
            //    to its own paragraph before the original paragraph.
            MoveTagsToSeparateParagraphs(doc);

            // 4. Build the report using LINQ Reporting Engine.
            ReportingEngine engine = new ReportingEngine();
            Order data = new Order(); // sample data
            engine.BuildReport(doc, data, "order");

            // 5. Save the final report.
            const string outputPath = "ReportResult.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }

        // Scans the document and moves any tag text (<<...>>) that is inside a run
        // to a separate paragraph placed before the original paragraph.
        private static void MoveTagsToSeparateParagraphs(Document doc)
        {
            // Collect paragraphs that need processing to avoid modifying the collection while iterating.
            List<Paragraph> paragraphs = new List<Paragraph>();
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
                paragraphs.Add(para);

            foreach (Paragraph para in paragraphs)
            {
                // Search runs for tag patterns.
                foreach (Run run in para.GetChildNodes(NodeType.Run, true))
                {
                    string text = run.Text;
                    int startIdx = text.IndexOf("<<", StringComparison.Ordinal);
                    int endIdx = text.IndexOf(">>", StringComparison.Ordinal);

                    // Simple detection of a tag inside the run.
                    if (startIdx >= 0 && endIdx > startIdx)
                    {
                        string tag = text.Substring(startIdx, endIdx - startIdx + 2);

                        // Remove the tag from the original run.
                        string newRunText = text.Remove(startIdx, tag.Length);
                        run.Text = newRunText;

                        // Insert a new paragraph before the current one containing only the tag.
                        Paragraph tagParagraph = (Paragraph)para.Clone(false);
                        Run tagRun = new Run(doc, tag);
                        tagParagraph.Runs.Clear();
                        tagParagraph.Runs.Add(tagRun);
                        para.ParentNode.InsertBefore(tagParagraph, para);
                        // Only handle the first tag per run for this example.
                        break;
                    }
                }
            }
        }
    }
}
