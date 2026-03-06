using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Tables;
using Aspose.Words.Notes; // Added for Footnote and FootnoteType

class CompareDocumentsDemo
{
    static void Main()
    {
        // Folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // ---------- Create the original document ----------
        Document docOriginal = new Document();                     // create blank document
        DocumentBuilder builder = new DocumentBuilder(docOriginal);

        // Paragraph text.
        builder.Writeln("Hello world! This is the first paragraph.");

        // Endnote.
        builder.InsertFootnote(FootnoteType.Endnote, "Original endnote text.");

        // Table with two cells.
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

        // ---------- Clone the document and make edits ----------
        Document docEdited = (Document)docOriginal.Clone(true);

        // Edit paragraph text.
        Paragraph firstParagraph = docEdited.FirstSection.Body.FirstParagraph;
        firstParagraph.Runs[0].Text = "hello world! this is the first paragraph, after editing.";
        firstParagraph.ParagraphFormat.Style = docEdited.Styles[StyleIdentifier.Heading1];

        // Edit endnote text.
        Footnote footnote = (Footnote)docEdited.GetChild(NodeType.Footnote, 0, true);
        footnote.FirstParagraph.Runs[0].Text = "Edited endnote text.";

        // Edit table cell.
        Table table = (Table)docEdited.GetChild(NodeType.Table, 0, true);
        table.FirstRow.Cells[1].FirstParagraph.Runs[0].Text = "Edited Cell 2 contents";

        // Edit textbox contents.
        Shape editedTextBox = (Shape)docEdited.GetChild(NodeType.Shape, 0, true);
        editedTextBox.FirstParagraph.Runs[0].Text = "Edited textbox contents";

        // Change DATE field property.
        ((FieldDate)docEdited.Range.Fields[0]).UseLunarCalendar = true;

        // Edit comment text.
        Comment editedComment = (Comment)docEdited.GetChild(NodeType.Comment, 0, true);
        editedComment.FirstParagraph.Runs[0].Text = "Edited comment.";

        // Edit header text.
        docEdited.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary]
            .FirstParagraph.Runs[0].Text = "Edited header contents.";

        // ---------- Compare documents with options ----------
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

        // Perform comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "John Doe", DateTime.Now, compareOptions);

        // Save the result with revisions.
        string outputPath = Path.Combine(artifactsDir, "Revision.CompareOptions.docx");
        docOriginal.Save(outputPath);
    }
}
