using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Topic
{
    // Initialise with empty strings to satisfy non‑nullable warnings.
    public string Title { get; set; } = string.Empty;
    public string Content { get; set; } = string.Empty;
}

public class BookmarkFromDotDemo
{
    public void Execute()
    {
        // Load the DOT template that contains a repeatable region for topics.
        // The template should have a syntax like:
        //   <<foreach [data.topics]>>
        //   <<[Title]>>
        //   <<[Content]>>
        //   <<end>>
        Document template = new Document("Template.dot");

        // Prepare a list of topics that will be merged into the template.
        List<Topic> topics = new List<Topic>
        {
            new Topic { Title = "Introduction", Content = "This is the introduction." },
            new Topic { Title = "Usage",        Content = "Details about usage." },
            new Topic { Title = "Conclusion",   Content = "Final remarks." }
        };

        // Build the report using the LINQ Reporting Engine.
        // The anonymous object supplies the data source; "data" is the name used in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, new { topics }, "data");

        // After the report is generated, insert a bookmark for each topic title.
        DocumentBuilder builder = new DocumentBuilder(template);
        foreach (Topic topic in topics)
        {
            // Locate the paragraph that contains the title text.
            Paragraph titleParagraph = template.GetChildNodes(NodeType.Paragraph, true)
                                               .Cast<Paragraph>()
                                               .FirstOrDefault(p => p.GetText().Contains(topic.Title));
            if (titleParagraph == null)
                continue; // Title not found – skip.

            // Create a bookmark name that is safe for Word (no spaces, unique).
            string bookmarkName = topic.Title.Replace(' ', '_');

            // Move the cursor to the start of the paragraph and start the bookmark.
            builder.MoveTo(titleParagraph);
            builder.StartBookmark(bookmarkName);

            // Move the cursor to the end of the same paragraph and close the bookmark.
            // The end of a paragraph is after its last child node.
            builder.MoveTo(titleParagraph.LastChild);
            builder.EndBookmark(bookmarkName);
        }

        // Save the final document. The format is inferred from the file extension.
        template.Save("Result.docx");
    }
}

public static class Program
{
    public static void Main()
    {
        new BookmarkFromDotDemo().Execute();
    }
}
