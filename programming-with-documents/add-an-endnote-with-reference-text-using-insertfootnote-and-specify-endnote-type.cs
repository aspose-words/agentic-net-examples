using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

namespace EndnoteExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some body text that will be referenced by the endnote.
            builder.Write("This is some sample text that will have an endnote.");

            // Insert an endnote with the specified reference text.
            builder.InsertFootnote(FootnoteType.Endnote, "This is the endnote content.");

            // Define the output file path (in the current working directory).
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "EndnoteExample.docx");

            // Save the document to the file system.
            doc.Save(outputPath);
        }
    }
}
