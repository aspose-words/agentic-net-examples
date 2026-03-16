using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;

namespace ExtractBetweenParagraphs
{
    // Visitor that extracts plain text between two paragraphs while skipping comments.
    class ExtractTextVisitor : DocumentVisitor
    {
        private readonly Paragraph _startParagraph;
        private readonly Paragraph _endParagraph;
        private bool _insideRange;
        private bool _insideComment;
        private readonly StringBuilder _builder = new StringBuilder();

        public ExtractTextVisitor(Paragraph startParagraph, Paragraph endParagraph)
        {
            _startParagraph = startParagraph;
            _endParagraph = endParagraph;
            _insideRange = false;
            _insideComment = false;
        }

        // Called when a paragraph starts.
        public override VisitorAction VisitParagraphStart(Paragraph paragraph)
        {
            // When we encounter the start paragraph, begin the range after it.
            if (paragraph == _startParagraph)
                _insideRange = true;

            // If the end paragraph is reached, finish the range before processing its content.
            if (paragraph == _endParagraph)
                _insideRange = false;

            return VisitorAction.Continue;
        }

        // Called when a run (piece of text) is encountered.
        public override VisitorAction VisitRun(Run run)
        {
            if (_insideRange && !_insideComment)
                _builder.Append(run.Text);
            return VisitorAction.Continue;
        }

        // Called when a paragraph ends – add a line break if we are still inside the range.
        public override VisitorAction VisitParagraphEnd(Paragraph paragraph)
        {
            if (_insideRange && !_insideComment)
                _builder.Append(Environment.NewLine);
            return VisitorAction.Continue;
        }

        // Comment range start – ignore everything inside the comment.
        public override VisitorAction VisitCommentRangeStart(CommentRangeStart commentRangeStart)
        {
            if (_insideRange)
                _insideComment = true;
            return VisitorAction.Continue;
        }

        // Comment range end – resume normal processing.
        public override VisitorAction VisitCommentRangeEnd(CommentRangeEnd commentRangeEnd)
        {
            if (_insideRange)
                _insideComment = false;
            return VisitorAction.Continue;
        }

        // Stand‑alone comment nodes – also ignored.
        public override VisitorAction VisitCommentStart(Comment comment)
        {
            if (_insideRange)
                _insideComment = true;
            return VisitorAction.Continue;
        }

        public override VisitorAction VisitCommentEnd(Comment comment)
        {
            if (_insideRange)
                _insideComment = false;
            return VisitorAction.Continue;
        }

        // Returns the collected text.
        public string GetText()
        {
            return _builder.ToString();
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the document.
            Document doc = new Document("Input.docx");

            // Identify the two paragraphs that bound the extraction range.
            // For this example we use the first two paragraphs that contain specific marker text.
            Paragraph startParagraph = null;
            Paragraph endParagraph = null;

            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                string txt = para.GetText().Trim();
                if (startParagraph == null && txt.Contains("[Start]"))
                    startParagraph = para;
                else if (startParagraph != null && txt.Contains("[End]"))
                {
                    endParagraph = para;
                    break;
                }
            }

            if (startParagraph == null || endParagraph == null)
                throw new InvalidOperationException("Start or end paragraph not found.");

            // Run the visitor over the whole document.
            var visitor = new ExtractTextVisitor(startParagraph, endParagraph);
            doc.Accept(visitor);

            // Extracted text without comments.
            string result = visitor.GetText();

            // Output or save the result.
            Console.WriteLine("Extracted text:");
            Console.WriteLine(result);
        }
    }
}
