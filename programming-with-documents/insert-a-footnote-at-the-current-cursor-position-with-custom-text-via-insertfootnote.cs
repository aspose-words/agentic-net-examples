using System;
using Aspose.Words;
using Aspose.Words.Notes;

namespace AsposeWordsFootnoteExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some text that will be referenced by the footnote.
            builder.Write("This sentence will have a footnote attached to it.");

            // Insert a footnote at the current cursor position with custom text.
            builder.InsertFootnote(FootnoteType.Footnote, "This is the custom footnote text.");

            // Save the document to the local file system.
            string outputPath = "FootnoteExample.docx";
            doc.Save(outputPath);
        }
    }
}
