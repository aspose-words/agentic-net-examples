using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 1. Add a title paragraph with Heading 1 style and centered alignment.
        // -------------------------------------------------
        builder.Writeln("Document Title");
        Paragraph titlePara = doc.FirstSection.Body.LastParagraph;
        titlePara.ParagraphFormat.StyleName = "Heading 1";
        titlePara.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // -------------------------------------------------
        // 2. Add a normal paragraph.
        // -------------------------------------------------
        builder.Writeln("This is the first paragraph of the document. It will demonstrate basic formatting.");

        // -------------------------------------------------
        // 3. Add a paragraph with custom indents and double line spacing.
        // -------------------------------------------------
        builder.Writeln("This paragraph has a left indent, a hanging indent, and double line spacing.");
        Paragraph formattedPara = doc.FirstSection.Body.LastParagraph;
        formattedPara.ParagraphFormat.LeftIndent = 20;            // points
        formattedPara.ParagraphFormat.FirstLineIndent = -20;      // hanging indent (negative first‑line indent)
        // Use Multiple rule with a factor of 2.0 to achieve double spacing.
        formattedPara.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        formattedPara.ParagraphFormat.LineSpacing = 2.0;

        // -------------------------------------------------
        // 4. Insert placeholder text that will be replaced later.
        // -------------------------------------------------
        builder.Writeln("Dear {CustomerName},");
        builder.Writeln("Thank you for your purchase.");

        // -------------------------------------------------
        // 5. Access the collection of all paragraphs in the body.
        // -------------------------------------------------
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;
        Console.WriteLine($"Total paragraphs after insertion: {paragraphs.Count}");

        // -------------------------------------------------
        // 6. Replace the placeholder with an actual name, preserving formatting.
        // -------------------------------------------------
        FindReplaceOptions replaceOptions = new FindReplaceOptions
        {
            MatchCase = false
        };
        doc.Range.Replace("{CustomerName}", "John Doe", replaceOptions);

        // -------------------------------------------------
        // 7. Insert a text box shape and write text inside it.
        // -------------------------------------------------
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
        textBox.TextBox.VerticalAnchor = TextBoxAnchor.Middle; // vertically center text
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("Important Notice");

        // -------------------------------------------------
        // 8. Save the document to a file.
        // -------------------------------------------------
        doc.Save("ManipulatedParagraphs.docx");
    }
}
