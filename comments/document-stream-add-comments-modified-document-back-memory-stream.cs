using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsDemo
{
    public static class DocumentCommentHelper
    {
        /// <summary>
        /// Loads a Word document from the provided input stream, adds a comment,
        /// and returns the modified document in a new memory stream.
        /// </summary>
        /// <param name="inputStream">Stream containing the source document.</param>
        /// <returns>MemoryStream with the modified document.</returns>
        public static MemoryStream AddCommentAndSave(Stream inputStream)
        {
            // Load the document from the input stream.
            Document doc = new Document(inputStream);

            // Use DocumentBuilder to edit the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // Create and attach a comment.
            Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
            comment.SetText("This is a sample comment added programmatically.");
            builder.CurrentParagraph.AppendChild(comment);

            // Save to a new memory stream.
            MemoryStream outputStream = new MemoryStream();
            doc.Save(outputStream, SaveFormat.Docx);
            outputStream.Position = 0;
            return outputStream;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: load a simple DOCX from embedded resources or a file.
            // For demonstration we create an empty document in memory.
            using (MemoryStream source = new MemoryStream())
            {
                // Create a blank document and save it to the source stream.
                Document emptyDoc = new Document();
                emptyDoc.Save(source, SaveFormat.Docx);
                source.Position = 0;

                // Call the helper to add a comment.
                MemoryStream result = DocumentCommentHelper.AddCommentAndSave(source);

                // Optionally write the result to a file to verify.
                using (FileStream file = new FileStream("Result.docx", FileMode.Create, FileAccess.Write))
                {
                    result.CopyTo(file);
                }
                Console.WriteLine("Document with comment saved as Result.docx");
            }
        }
    }
}
