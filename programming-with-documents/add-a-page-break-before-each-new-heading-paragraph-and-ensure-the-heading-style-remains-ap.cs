using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsPageBreakExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add sample content with heading styles.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
            builder.Writeln("Document Title");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This is some body text for chapter 1.");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("More detailed text.");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Body text for chapter 2.");

            // Insert a page break before each heading paragraph while preserving the heading style.
            DocumentBuilder insertBuilder = new DocumentBuilder(doc);
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                StyleIdentifier styleId = para.ParagraphFormat.StyleIdentifier;
                if (styleId >= StyleIdentifier.Heading1 && styleId <= StyleIdentifier.Heading9)
                {
                    insertBuilder.MoveTo(para);
                    insertBuilder.InsertBreak(BreakType.PageBreak);
                }
            }

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
            doc.Save(outputPath);
        }
    }
}
