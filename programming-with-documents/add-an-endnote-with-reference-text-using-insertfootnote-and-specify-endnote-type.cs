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

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some text that will be referenced by the endnote.
            builder.Write("This sentence will have an endnote attached.");

            // Insert an endnote with the desired reference text.
            builder.InsertFootnote(FootnoteType.Endnote, "This is the endnote text.");

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document to a file.
            string outputPath = Path.Combine(outputDir, "EndnoteExample.docx");
            doc.Save(outputPath);
        }
    }
}
