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
        public string Name { get; set; }      // Bookmark name.
        public string Content { get; set; }   // Text to place inside the bookmark.
    }

    public class Program
    {
        public static void Main()
        {
            // Load the DOC template that may contain other merge fields.
            Document doc = new Document("Template.docx");

            // Example data source – could be retrieved from a database or any LINQ query.
            List<Topic> topics = GetTopicsFromDataSource();

            // If the template contains other merge fields, populate them using ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // Assuming the template has a placeholder for a title field.
            var reportData = new { Title = "Dynamic Topics Report" };
            engine.BuildReport(doc, reportData, "data");

            // Insert bookmarks dynamically at the end of the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            foreach (Topic topic in topics)
            {
                // Start a new paragraph for each topic.
                builder.Writeln();

                // Insert a bookmark with the topic name.
                builder.StartBookmark(topic.Name);
                builder.Write(topic.Content);
                builder.EndBookmark(topic.Name);
            }

            // Save the resulting document.
            doc.Save("Result.docx");
        }

        // Mock method that returns a list of topics using LINQ.
        private static List<Topic> GetTopicsFromDataSource()
        {
            // Sample raw data.
            var rawData = new[]
            {
                new { Id = 1, Title = "Introduction", Body = "This is the introduction." },
                new { Id = 2, Title = "Usage", Body = "Details on how to use the product." },
                new { Id = 3, Title = "Conclusion", Body = "Final thoughts and summary." }
            };

            // LINQ projection to the Topic model.
            return rawData
                .Select(item => new Topic
                {
                    Name = $"Bookmark_{item.Id}",   // Ensure unique bookmark names.
                    Content = $"{item.Title}: {item.Body}"
                })
                .ToList();
        }
    }
}
