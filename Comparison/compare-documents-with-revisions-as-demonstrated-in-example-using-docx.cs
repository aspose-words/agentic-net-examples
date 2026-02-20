using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Notes;      // Added for Footnote
using Aspose.Words.Tables;     // Added for Table

class CompareDocumentsWithRevisions
{
    static void Main()
    {
        // Create the original document and populate it with various elements.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);

        // Paragraph text.
        builder.Writeln("Hello world! This is the first paragraph.");

        // Endnote.
        builder.InsertFootnote(FootnoteType.Endnote, "Original endnote text.");

        // Table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Original cell 1 text");
        builder.InsertCell();
        builder.Write("Original cell 2 text");
        builder.EndTable();

        // Textbox.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 20);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("Original textbox contents");

        // DATE field.
        builder.MoveTo(docOriginal.FirstSection.Body.AppendParagraph(""));
        builder.InsertField(" DATE ");

        // Comment.
        Comment comment = new Comment(docOriginal, "John Doe", "J.D.", DateTime.Now);
        comment.SetText("Original comment.");
        builder.CurrentParagraph.AppendChild(comment);

        // Header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Original header contents.");

        // Clone the document and apply edits.
        Document docEdited = (Document)docOriginal.Clone(true);
        Paragraph firstParagraph = docEdited.FirstSection.Body.FirstParagraph;
        firstParagraph.Runs[0].Text = "hello world! this is the first paragraph, after editing.";
        firstParagraph.ParagraphFormat.Style = docEdited.Styles[StyleIdentifier.Heading1];
        ((Footnote)docEdited.GetChild(NodeType.Footnote, 0, true)).FirstParagraph.Runs[1].Text = "Edited endnote text.";
        ((Table)docEdited.GetChild(NodeType.Table, 0, true)).FirstRow.Cells[1].FirstParagraph.Runs[0].Text = "Edited Cell 2 contents";
        ((Shape)docEdited.GetChild(NodeType.Shape, 0, true)).FirstParagraph.Runs[0].Text = "Edited textbox contents";
        ((FieldDate)docEdited.Range.Fields[0]).UseLunarCalendar = true;
        ((Comment)docEdited.GetChild(NodeType.Comment, 0, true)).FirstParagraph.Runs[0].Text = "Edited comment.";
        docEdited.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].FirstParagraph.Runs[0].Text = "Edited header contents.";

        // Configure comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = false,
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison – revisions are added to docOriginal.
        docOriginal.Compare(docEdited, "John Doe", DateTime.Now, compareOptions);

        // Save the resulting document with revisions.
        docOriginal.Save("Revision.CompareOptions.docx");
    }
}
