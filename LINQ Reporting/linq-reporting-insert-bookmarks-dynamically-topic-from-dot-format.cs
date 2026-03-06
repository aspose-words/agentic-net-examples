using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load a DOT template that contains LINQ Reporting placeholders.
        // The constructor Document(string) is the approved load rule.
        Document doc = new Document("Template.dot");

        // Prepare a data source – a list of topics that will be merged into the template.
        // Each topic has a Title and Content that can be referenced in the template as <<[topics.Title]>> etc.
        List<Topic> topics = new List<Topic>
        {
            new Topic { Title = "Introduction", Content = "Welcome to the report." },
            new Topic { Title = "Analysis",     Content = "Data analysis goes here." },
            new Topic { Title = "Conclusion",   Content = "Thanks for reading." }
        };

        // Populate the template using the LINQ Reporting Engine.
        // BuildReport(Document, object[], string[]) is the appropriate rule for multiple data sources.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, new object[] { topics }, new string[] { "topics" });

        // After the report is built, insert a bookmark for each topic.
        // The bookmarks are created dynamically based on the topic titles.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd(); // Position the cursor at the end of the document.

        foreach (Topic topic in topics)
        {
            // Convert the title to a valid bookmark name (letters, digits, underscore; starts with a letter).
            string bookmarkName = MakeValidBookmarkName(topic.Title);

            // Start the bookmark.
            builder.StartBookmark(bookmarkName);

            // Insert the title as a heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln(topic.Title);

            // Insert the content as normal text.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln(topic.Content);

            // End the bookmark.
            builder.EndBookmark(bookmarkName);
        }

        // Save the final document.
        // The Save(string) method is the approved save rule.
        doc.Save("Report.docx");
    }

    // Helper method to ensure the bookmark name complies with Word's naming rules.
    static string MakeValidBookmarkName(string title)
    {
        StringBuilder sb = new StringBuilder();
        foreach (char ch in title)
        {
            if (char.IsLetterOrDigit(ch) || ch == '_')
                sb.Append(ch);
            else if (char.IsWhiteSpace(ch))
                sb.Append('_');
        }

        // Ensure the name starts with a letter; prepend 'B' if necessary.
        if (sb.Length == 0 || !char.IsLetter(sb[0]))
            sb.Insert(0, 'B');

        return sb.ToString();
    }

    // Simple POCO representing a topic in the report.
    public class Topic
    {
        public string Title { get; set; }
        public string Content { get; set; }
    }
}
