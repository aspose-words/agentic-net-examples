using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace AsposeWordsBookmarkDemo
{
    // Simple data model used as a data source for the reporting engine.
    public class Topic
    {
        public string Title { get; set; }
        public string Content { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Load the DOC template that contains a placeholder for each topic.
            //    The placeholder can be any unique text, e.g. "<<Topic>>".
            Document doc = new Document("Template.docx");

            // 2. Prepare a list of topics that will be merged into the template.
            List<Topic> topics = new List<Topic>
            {
                new Topic { Title = "Introduction", Content = "This is the introduction." },
                new Topic { Title = "Usage", Content = "How to use the product." },
                new Topic { Title = "Conclusion", Content = "Final remarks." }
            };

            // 3. Use the ReportingEngine to populate the template.
            //    The template can reference the data source members via <<[topics.Title]>> etc.
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("topics") must match the name used in the template.
            engine.BuildReport(doc, topics, "topics");

            // 4. After the report is built, insert a bookmark for each topic dynamically.
            //    We locate the paragraphs that contain the topic title (populated by the engine)
            //    and wrap the title text with a bookmark whose name is derived from the title.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Get all paragraphs in the document.
            NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            foreach (Paragraph para in allParagraphs.OfType<Paragraph>())
            {
                // Trim the paragraph text to ignore trailing paragraph marks.
                string paraText = para.GetText().Trim();

                // Find a topic whose title exactly matches the paragraph text.
                Topic matchingTopic = topics.FirstOrDefault(t => t.Title.Equals(paraText, StringComparison.OrdinalIgnoreCase));
                if (matchingTopic == null) continue; // No matching topic – skip.

                // Move the builder cursor to the start of the paragraph.
                builder.MoveTo(para.FirstChild);

                // Define a bookmark name (Word bookmark names cannot contain spaces).
                string bookmarkName = "bm_" + matchingTopic.Title.Replace(' ', '_');

                // Insert the bookmark start, write the title (already present, but we rewrite to be safe),
                // and then insert the bookmark end.
                builder.StartBookmark(bookmarkName);
                // The title text is already in the paragraph; we simply preserve it.
                builder.Write(matchingTopic.Title);
                builder.EndBookmark(bookmarkName);
            }

            // 5. Save the resulting document.
            doc.Save("ReportWithBookmarks.docx");
        }
    }
}
