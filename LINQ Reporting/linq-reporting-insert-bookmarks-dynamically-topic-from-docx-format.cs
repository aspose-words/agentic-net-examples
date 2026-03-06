using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsBookmarkDemo
{
    // Simple POCO representing a topic to be inserted into the document.
    public class Topic
    {
        public string Title { get; set; }
        public string Content { get; set; }
    }

    // Wrapper class used as a data source for the ReportingEngine.
    public class ReportDataSource
    {
        public List<Topic> Topics { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Prepare sample data.
            var topics = new List<Topic>
            {
                new Topic { Title = "Introduction", Content = "This is the introduction section." },
                new Topic { Title = "Usage", Content = "Details about how to use the product." },
                new Topic { Title = "Conclusion", Content = "Final thoughts and summary." }
            };

            // Load the DOCX template that contains the reporting tags.
            // The template should have a foreach block like:
            // <<foreach [Topics]>><<[Title]>>\n<<[Content]>>\n<</foreach>>
            Document doc = new Document("Template.docx");

            // Populate the template using the ReportingEngine.
            var dataSource = new ReportDataSource { Topics = topics };
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "data");

            // After the report is built, insert a bookmark for each topic title.
            DocumentBuilder builder = new DocumentBuilder(doc);

            foreach (var topic in topics)
            {
                // Find the first paragraph that contains the exact title text.
                // LINQ is used to locate the paragraph node.
                var paragraph = doc.GetChildNodes(Aspose.Words.NodeType.Paragraph, true)
                                  .Cast<Aspose.Words.Paragraph>()
                                  .FirstOrDefault(p => p.GetText().Trim() == topic.Title);

                if (paragraph != null)
                {
                    // Move the cursor to the start of the paragraph.
                    builder.MoveTo(paragraph.FirstChild);

                    // Create a bookmark with the same name as the title.
                    // Bookmark names must be unique; if needed, adjust the name.
                    builder.StartBookmark(topic.Title);
                    // The bookmark range is the whole paragraph.
                    builder.EndBookmark(topic.Title);
                }
            }

            // Save the resulting document.
            doc.Save("ReportWithBookmarks.docx");
        }
    }
}
