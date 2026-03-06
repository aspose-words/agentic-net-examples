using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCM template.
        Document doc = new Document("Template.docm");

        // Sample data source: a list of topics.
        var topics = new List<Topic>
        {
            new Topic { Title = "Introduction", Content = "This is the introduction." },
            new Topic { Title = "Usage",        Content = "How to use the product." },
            new Topic { Title = "Conclusion",   Content = "Final remarks and summary." }
        };

        // Populate the template with the data source using ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The data source name ("topics") can be referenced in the template if needed.
        engine.BuildReport(doc, topics, "topics");

        // Insert bookmarks dynamically for each topic.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Position the builder at the end of the document to start inserting.
        builder.MoveToDocumentEnd();

        foreach (var t in topics)
        {
            // Write the topic title as a separate paragraph.
            builder.Writeln(t.Title);

            // Create a valid bookmark name from the title.
            string bookmarkName = MakeValidBookmarkName(t.Title);

            // Start the bookmark before the content.
            builder.StartBookmark(bookmarkName);
            // Write the topic content.
            builder.Writeln(t.Content);
            // End the bookmark after the content.
            builder.EndBookmark(bookmarkName);
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }

    // Helper method to generate a bookmark name that complies with Word's rules.
    static string MakeValidBookmarkName(string title)
    {
        // Keep only letters, digits and underscores.
        string name = new string(title.Where(c => char.IsLetterOrDigit(c) || c == '_').ToArray());

        // Ensure the name is not empty.
        if (string.IsNullOrEmpty(name))
            name = "Bookmark";

        // Word limits bookmark names to 40 characters.
        if (name.Length > 40)
            name = name.Substring(0, 40);

        return name;
    }

    // Simple POCO representing a topic.
    public class Topic
    {
        public string Title { get; set; }
        public string Content { get; set; }
    }
}
