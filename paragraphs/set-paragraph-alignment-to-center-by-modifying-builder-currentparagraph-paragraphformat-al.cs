using System;
using Aspose.Words;

namespace AsposeWordsParagraphAlignment
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a paragraph of text.
            builder.Writeln("This paragraph will be centered.");

            // Set the alignment of the current paragraph to center.
            builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Save the document.
            string outputPath = "ParagraphAlignmentCenter.docx";
            doc.Save(outputPath);
        }
    }
}
