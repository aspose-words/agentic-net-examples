using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace StyledContentExtraction
{
    public class Program
    {
        public static void Main()
        {
            // -------------------------------------------------
            // 1. Create a sample source document with styled paragraphs.
            // -------------------------------------------------
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            // Normal paragraph.
            builder.Writeln("Introduction paragraph.");

            // Heading 1 – start marker.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Start of Styled Section");

            // Normal paragraphs inside the range.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("First content paragraph.");
            builder.Writeln("Second content paragraph.");

            // Heading 2 – end marker.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("End of Styled Section");

            // Normal paragraph after the range.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Conclusion paragraph.");

            // Save the source document locally.
            const string sourcePath = "source.docx";
            sourceDoc.Save(sourcePath);

            // -------------------------------------------------
            // 2. Load the document for extraction.
            // -------------------------------------------------
            Document loadedDoc = new Document(sourcePath);
            Body body = loadedDoc.FirstSection.Body;

            // -------------------------------------------------
            // 3. Locate start and end paragraphs based on styles.
            // -------------------------------------------------
            Paragraph startParagraph = null;
            Paragraph endParagraph = null;

            foreach (Paragraph para in body.Paragraphs)
            {
                StyleIdentifier styleId = para.ParagraphFormat.StyleIdentifier;
                if (styleId == StyleIdentifier.Heading1 && startParagraph == null)
                {
                    startParagraph = para;
                }
                else if (styleId == StyleIdentifier.Heading2 && startParagraph != null && endParagraph == null)
                {
                    endParagraph = para;
                    break; // End found.
                }
            }

            if (startParagraph == null)
                throw new InvalidOperationException("Start paragraph with Heading 1 style not found.");
            if (endParagraph == null)
                throw new InvalidOperationException("End paragraph with Heading 2 style not found.");

            // -------------------------------------------------
            // 4. Build a new document containing the extracted range.
            // -------------------------------------------------
            Document resultDoc = new Document();
            resultDoc.RemoveAllChildren();

            Section resultSection = new Section(resultDoc);
            resultDoc.AppendChild(resultSection);

            Body resultBody = new Body(resultDoc);
            resultSection.AppendChild(resultBody);

            bool copy = false;
            foreach (Paragraph para in body.Paragraphs)
            {
                if (para == startParagraph)
                    copy = true;

                if (copy)
                {
                    // Import the paragraph into the destination document to avoid cross‑document node errors.
                    Node importedNode = resultDoc.ImportNode(para, true);
                    resultBody.AppendChild(importedNode);
                }

                if (para == endParagraph)
                    break;
            }

            // -------------------------------------------------
            // 5. Save the extracted content.
            // -------------------------------------------------
            const string resultPath = "extracted.docx";
            resultDoc.Save(resultPath);

            // Verify that the output file was created.
            if (!File.Exists(resultPath))
                throw new InvalidOperationException("Extraction output file was not created.");

            Console.WriteLine("Styled content segment extracted successfully.");
        }
    }
}
