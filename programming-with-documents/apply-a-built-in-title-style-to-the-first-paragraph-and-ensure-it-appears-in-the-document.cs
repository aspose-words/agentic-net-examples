using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing.Charts;

namespace AsposeWordsTitleStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some text for the first paragraph.
            builder.Writeln("Document Title");

            // Apply the built‑in "Title" style to the first paragraph.
            Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
            firstParagraph.ParagraphFormat.StyleName = "Title";

            // Ensure the paragraph appears in the document outline by setting its outline level.
            firstParagraph.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;

            // Save the document to the local file system.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TitleStyle.docx");
            doc.Save(outputPath);
        }
    }
}
