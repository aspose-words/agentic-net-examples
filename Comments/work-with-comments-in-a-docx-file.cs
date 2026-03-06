using System;
using System.IO;
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

            // Use DocumentBuilder to add some text.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is the first paragraph.");
            builder.Writeln("This is the second paragraph where we will add a comment.");

            // Create a top‑level comment.
            Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
            comment.SetText("Please review this sentence.");

            // Insert the comment range start, the commented text, and the comment range end.
            // The comment will be anchored to the text "second paragraph".
            Paragraph para = doc.FirstSection.Body.Paragraphs[1]; // second paragraph
            // Insert range start.
            para.AppendChild(new CommentRangeStart(doc, comment.Id));
            // Insert the text that will be commented.
            para.AppendChild(new Run(doc, "This is the second paragraph where we will add a comment."));
            // Insert range end.
            para.AppendChild(new CommentRangeEnd(doc, comment.Id));
            // Append the comment itself after the range.
            para.AppendChild(comment);

            // Add a reply to the comment.
            comment.AddReply("Jane Smith", "JS", DateTime.Now, "I have reviewed it and it looks good.");

            // Save the document to a DOCX file.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "CommentsDemo.docx");
            doc.Save(outputPath);

            // Load the saved document and iterate over all top‑level comments.
            Document loadedDoc = new Document(outputPath);
            NodeCollection commentNodes = loadedDoc.GetChildNodes(NodeType.Comment, true);

            Console.WriteLine("Comments in the document:");
            foreach (Comment topLevelComment in commentNodes)
            {
                // A top‑level comment has no Comment as its parent.
                if (!(topLevelComment.ParentNode is Comment))
                {
                    Console.WriteLine($"Comment ID: {topLevelComment.Id}");
                    Console.WriteLine($"Author: {topLevelComment.Author}");
                    Console.WriteLine($"Date: {topLevelComment.DateTime}");
                    Console.WriteLine($"Text: {topLevelComment.GetText().Trim()}");

                    // List any replies.
                    foreach (Comment reply in topLevelComment.Replies)
                    {
                        Console.WriteLine($"\tReply by {reply.Author} on {reply.DateTime}");
                        Console.WriteLine($"\tReply text: {reply.GetText().Trim()}");
                    }
                }
            }
        }
    }
}
