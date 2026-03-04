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
            // Load the DOCM template that contains the report placeholders.
            // The template can have fields like <<foreach [topics]>> <<[Title]>> <<[Content]>> <</foreach>>
            Document doc = new Document("Template.docm");

            // Prepare a collection of topics that will be used as the data source.
            List<Topic> topics = new List<Topic>
            {
                new Topic { Title = "Introduction", Content = "This is the introduction." },
                new Topic { Title = "Usage", Content = "Details about usage." },
                new Topic { Title = "Conclusion", Content = "Final thoughts." }
            };

            // Use Aspose.Words ReportingEngine to populate the template with the topics collection.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "topics" must match the name used in the template's foreach tag.
            engine.BuildReport(doc, new object[] { topics }, new[] { "topics" });

            // After the report is built, insert a bookmark for each topic dynamically.
            // The bookmarks will be placed at the start of each topic title.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Find all paragraphs that contain a topic title (simple LINQ query on paragraphs).
            var titleParagraphs = doc.FirstSection.Body.Paragraphs
                .Cast<Paragraph>()
                .Where(p => topics.Any(t => p.GetText().Contains(t.Title)))
                .ToList();

            int index = 0;
            foreach (var paragraph in titleParagraphs)
            {
                // Move the builder cursor to the paragraph that holds the title.
                builder.MoveTo(paragraph);

                // Define a unique bookmark name for each topic.
                string bookmarkName = $"Topic_{index}";

                // Insert the bookmark start, write the title (already present), then insert the end.
                builder.StartBookmark(bookmarkName);
                // The title text is already in the paragraph, so we just close the bookmark.
                builder.EndBookmark(bookmarkName);

                index++;
            }

            // Save the resulting document. The format is inferred from the extension.
            doc.Save("ReportWithBookmarks.docx");
        }
    }
}
