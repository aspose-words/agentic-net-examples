using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsParagraphStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Apply the built‑in Heading1 style to the current paragraph.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            // Add some text – this paragraph will be formatted as Heading1.
            builder.Writeln("Sample Heading 1");

            // Save the document to the local file system.
            doc.Save("Output.docx");
        }
    }
}
