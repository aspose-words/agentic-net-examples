using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace AsposeWordsCommentsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for easy content insertion.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some initial text that will be the anchor for the comment.
            builder.Write("Hello world! ");

            // Create a top‑level comment with author information.
            Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
            comment.SetText("This is a top‑level comment.");

            // Insert a comment range that surrounds the text we want to comment.
            // The range start and end must have the same Id as the comment.
            builder.CurrentParagraph.AppendChild(new CommentRangeStart(doc, comment.Id));
            builder.Write("Commented text.");
            builder.CurrentParagraph.AppendChild(new CommentRangeEnd(doc, comment.Id));

            // Append the comment itself to the paragraph (it will appear in the margin).
            builder.CurrentParagraph.AppendChild(comment);

            // Add a reply to the top‑level comment.
            comment.AddReply("Jane Smith", "JS", DateTime.Now, "This is a reply to the comment.");

            // -----------------------------------------------------------------
            // OPTIONAL: Iterate over all comments and print their details to console.
            // -----------------------------------------------------------------
            foreach (Comment topLevelComment in doc.GetChildNodes(NodeType.Comment, true))
            {
                // Only process top‑level comments (those without an ancestor comment).
                if (topLevelComment.Ancestor == null)
                {
                    Console.WriteLine($"Comment ID: {topLevelComment.Id}");
                    Console.WriteLine($"Author: {topLevelComment.Author}");
                    Console.WriteLine($"Date: {topLevelComment.DateTime}");
                    Console.WriteLine($"Text: {topLevelComment.GetText().Trim()}");
                    Console.WriteLine();

                    // Print any replies.
                    foreach (Comment reply in topLevelComment.Replies)
                    {
                        Console.WriteLine($"\tReply ID: {reply.Id}");
                        Console.WriteLine($"\tAuthor: {reply.Author}");
                        Console.WriteLine($"\tDate: {reply.DateTime}");
                        Console.WriteLine($"\tText: {reply.GetText().Trim()}");
                        Console.WriteLine();
                    }
                }
            }

            // Save the document to a DOCX file.
            doc.Save("CommentsDemo.docx");
        }
    }
}
