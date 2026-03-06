using System;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsParagraphDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to insert three paragraphs.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First paragraph – will be styled as Heading 1 and centered.
            builder.Writeln("Chapter 1: Introduction");
            Paragraph headingParagraph = doc.FirstSection.Body.Paragraphs[0];
            headingParagraph.ParagraphFormat.StyleName = "Heading 1";
            headingParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Second paragraph – normal text.
            builder.Writeln("This is the first paragraph of the chapter. It contains some sample text.");

            // Third paragraph – will be right aligned after a find-and-replace operation.
            builder.Writeln("Please review the document. The word 'review' will be highlighted later.");

            // Demonstrate accessing paragraphs via the ParagraphCollection.
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            // Change the line spacing of the second paragraph.
            Paragraph secondParagraph = paragraphs[1];
            secondParagraph.ParagraphFormat.LineSpacing = 18; // points
            secondParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;

            // Perform a find-and-replace that also changes paragraph formatting.
            FindReplaceOptions replaceOptions = new FindReplaceOptions();
            // Apply right alignment to any paragraph that contains the word "review".
            replaceOptions.ApplyParagraphFormat.Alignment = ParagraphAlignment.Right;

            // Replace the word "review" with "reviewed".
            doc.Range.Replace("review", "reviewed", replaceOptions);

            // Verify that the third paragraph is now right aligned.
            Paragraph thirdParagraph = paragraphs[2];
            Console.WriteLine("Third paragraph alignment after replace: " + thirdParagraph.ParagraphFormat.Alignment);

            // Save the document to disk.
            doc.Save("ParagraphDemo.docx");
        }
    }
}
