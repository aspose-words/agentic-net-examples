using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace AsposeWordsBookmarkDemo
{
    // Simple data model representing a topic.
    public class Topic
    {
        public string Title { get; set; }
        public string Content { get; set; }

        // Constructor to satisfy non‑nullable warnings.
        public Topic(string title, string content)
        {
            Title = title;
            Content = content;
        }
    }

    public class Program
    {
        // Entry point.
        public static void Main()
        {
            // Load the DOCX template that contains the reporting placeholders.
            Document template = new Document("Template.docx");

            // Sample data source – a list of topics.
            List<Topic> topics = new List<Topic>
            {
                new Topic("Introduction", "Welcome to the introduction."),
                new Topic("Usage", "How to use the product."),
                new Topic("Conclusion", "Final thoughts.")
            };

            // Build the report using Aspose.Words ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The template can reference the data source as "topics".
            engine.BuildReport(template, topics, "topics");

            // After the report is built, insert bookmarks for each topic.
            // Use LINQ to enumerate the topics together with their index.
            DocumentBuilder builder = new DocumentBuilder(template);

            // Get all paragraphs in the document (including those inside tables, headers, etc.).
            var allParagraphs = template.GetChildNodes(NodeType.Paragraph, true)
                                         .Cast<Paragraph>()
                                         .ToList();

            foreach (var item in topics.Select((t, i) => new { Topic = t, Index = i }))
            {
                // Find the paragraph that contains the topic title.
                Paragraph titleParagraph = allParagraphs
                    .FirstOrDefault(p => p.GetText().Trim().Contains(item.Topic.Title));

                if (titleParagraph != null)
                {
                    // Move the builder cursor to the found paragraph.
                    builder.MoveTo(titleParagraph);

                    // Insert a bookmark that spans the whole paragraph.
                    string bookmarkName = $"Topic_{item.Index + 1}";
                    builder.StartBookmark(bookmarkName);
                    // The cursor is already positioned at the start of the paragraph; the bookmark will cover it.
                    builder.EndBookmark(bookmarkName);
                }
            }

            // Save the final document.
            template.Save("ReportWithBookmarks.docx");
        }
    }
}
