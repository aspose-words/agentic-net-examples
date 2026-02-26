using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsBookmarkDemo
{
    // Simple data model representing a topic.
    public class Topic
    {
        public string Title { get; set; }
        public string Content { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCM template. The template can be an empty document or contain static content.
            Document doc = new Document(@"C:\Templates\ReportTemplate.docm");

            // Example source data – in a real scenario this could come from a database, JSON, etc.
            List<Topic> allTopics = new List<Topic>
            {
                new Topic { Title = "Introduction", Content = "This is the introduction." },
                new Topic { Title = "Usage", Content = "How to use the product." },
                new Topic { Title = "Conclusion", Content = "Final thoughts." }
            };

            // Use LINQ to filter or order topics as needed.
            // For demonstration we order them alphabetically by title.
            var orderedTopics = allTopics.OrderBy(t => t.Title).ToList();

            // Use DocumentBuilder to insert each topic as a separate paragraph with a bookmark.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the cursor to the end of the document (or any desired location).
            builder.MoveToDocumentEnd();

            for (int i = 0; i < orderedTopics.Count; i++)
            {
                Topic topic = orderedTopics[i];
                string bookmarkName = $"Topic_{i + 1}";

                // Insert a paragraph for the title.
                builder.Writeln(topic.Title);
                // Mark the start of the bookmark.
                builder.StartBookmark(bookmarkName);
                // Insert the content that belongs to the bookmark.
                builder.Writeln(topic.Content);
                // Mark the end of the bookmark.
                builder.EndBookmark(bookmarkName);
            }

            // Optionally, use ReportingEngine to process any remaining merge fields in the template.
            // (If the template contains <<[DataSource.Property]>> style fields.)
            ReportingEngine engine = new ReportingEngine();
            // BuildReport can be called with a data source; here we pass the list of topics.
            engine.BuildReport(doc, orderedTopics, "Topics");

            // Save the resulting document. The format is inferred from the extension.
            doc.Save(@"C:\Output\GeneratedReport.docx");
        }
    }
}
