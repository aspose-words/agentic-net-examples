using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace AsposeWordsParagraphFormatting
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set paragraph alignment to Justify.
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;

            // Enable word wrap (wrap by whole words) for the paragraph.
            builder.ParagraphFormat.WordWrap = true;

            // Write a sample paragraph that will demonstrate the justification and word wrap.
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                            "Sed non risus. Suspendisse lectus tortor, dignissim sit amet, " +
                            "adipiscing nec, ultricies sed, dolor. Cras elementum ultrices diam. " +
                            "Maecenas ligula massa, varius a, semper congue, euismod non, mi.");

            // Save the document to the local file system.
            string outputPath = "JustifiedParagraph.docx";
            doc.Save(outputPath);
        }
    }
}
