using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;
using Aspose.Words.Fields; // Added for FieldType enum

namespace AsposeWordsParagraphDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a DocumentBuilder which will be used to add and edit content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ------------------------------------------------------------
            // 1. Work with headers and footers.
            // ------------------------------------------------------------
            // Move the cursor to the primary header of the first section and add text.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Document Header - Page ");

            // Insert a field that will display the current page number.
            builder.InsertField(FieldType.FieldPage, true);

            // Move the cursor to the primary footer and add text.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write("Document Footer - Confidential");

            // Return the cursor to the main body of the document.
            builder.MoveToDocumentEnd();

            // ------------------------------------------------------------
            // 2. Insert paragraphs and a footnote.
            // ------------------------------------------------------------
            builder.Writeln("First paragraph of the document.");
            builder.Writeln("Second paragraph contains a footnote reference.");

            // Insert a footnote after the current paragraph.
            Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "This is the footnote text.");

            // Move the builder inside the footnote to add more content.
            builder.MoveTo(footnote.FirstParagraph);
            builder.Write(" Additional text inside the footnote.");

            // Return to the end of the main document body.
            builder.MoveToDocumentEnd();

            // ------------------------------------------------------------
            // 3. Use the document's Range to perform a find-and-replace.
            // ------------------------------------------------------------
            // Replace the word "First" with "1st" throughout the whole document.
            int replaceCount = doc.Range.Replace("First", "1st");
            Console.WriteLine($"Replacements made: {replaceCount}");

            // ------------------------------------------------------------
            // 4. Extract content between two paragraphs using their ranges.
            // ------------------------------------------------------------
            // Get references to the first and second paragraphs.
            Paragraph firstPara = doc.FirstSection.Body.Paragraphs[0];
            Paragraph secondPara = doc.FirstSection.Body.Paragraphs[1];

            // The text of a paragraph's Range includes the paragraph break.
            string firstParaText = firstPara.Range.Text.TrimEnd('\r', '\a');
            string secondParaText = secondPara.Range.Text.TrimEnd('\r', '\a');

            Console.WriteLine("First paragraph text: " + firstParaText);
            Console.WriteLine("Second paragraph text: " + secondParaText);

            // ------------------------------------------------------------
            // 5. Demonstrate extracting the text of the footnote separator.
            // ------------------------------------------------------------
            // The footnote separator is a special story that separates footnotes from the main text.
            FootnoteSeparator separator = doc.FootnoteSeparators[FootnoteSeparatorType.FootnoteSeparator];
            string separatorText = separator.Range.Text.Trim();
            Console.WriteLine("Footnote separator text: '" + separatorText + "'");

            // ------------------------------------------------------------
            // 6. Save the document to a DOCX file.
            // ------------------------------------------------------------
            doc.Save("ParagraphsHeadersFootnotes.docx");
        }
    }
}
