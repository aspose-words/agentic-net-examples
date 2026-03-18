using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;

class TagValidatorVisitor : DocumentVisitor
{
    private readonly Stack<StructuredDocumentTag> _sdtStack = new Stack<StructuredDocumentTag>();
    private readonly Stack<BookmarkStart> _bookmarkStack = new Stack<BookmarkStart>();
    private readonly Stack<CommentRangeStart> _commentStack = new Stack<CommentRangeStart>();

    public override VisitorAction VisitStructuredDocumentTagStart(StructuredDocumentTag sdt)
    {
        _sdtStack.Push(sdt);
        return VisitorAction.Continue;
    }

    public override VisitorAction VisitStructuredDocumentTagEnd(StructuredDocumentTag sdt)
    {
        if (_sdtStack.Count == 0)
            throw new InvalidOperationException("StructuredDocumentTag end without matching start.");

        var start = _sdtStack.Pop();
        if (start.Id != sdt.Id)
            throw new InvalidOperationException($"Mismatched StructuredDocumentTag IDs: start {start.Id}, end {sdt.Id}.");

        return VisitorAction.Continue;
    }

    public override VisitorAction VisitBookmarkStart(BookmarkStart bookmarkStart)
    {
        _bookmarkStack.Push(bookmarkStart);
        return VisitorAction.Continue;
    }

    public override VisitorAction VisitBookmarkEnd(BookmarkEnd bookmarkEnd)
    {
        if (_bookmarkStack.Count == 0)
            throw new InvalidOperationException("Bookmark end without matching start.");

        var start = _bookmarkStack.Pop();
        if (start.Name != bookmarkEnd.Name)
            throw new InvalidOperationException($"Mismatched Bookmark names: start {start.Name}, end {bookmarkEnd.Name}.");

        return VisitorAction.Continue;
    }

    public override VisitorAction VisitCommentRangeStart(CommentRangeStart commentStart)
    {
        _commentStack.Push(commentStart);
        return VisitorAction.Continue;
    }

    public override VisitorAction VisitCommentRangeEnd(CommentRangeEnd commentEnd)
    {
        if (_commentStack.Count == 0)
            throw new InvalidOperationException("CommentRange end without matching start.");

        var start = _commentStack.Pop();
        if (start.Id != commentEnd.Id)
            throw new InvalidOperationException($"Mismatched Comment IDs: start {start.Id}, end {commentEnd.Id}.");

        return VisitorAction.Continue;
    }

    public void Verify()
    {
        if (_sdtStack.Count > 0)
            throw new InvalidOperationException("Unclosed StructuredDocumentTag(s) remain.");

        if (_bookmarkStack.Count > 0)
            throw new InvalidOperationException("Unclosed Bookmark(s) remain.");

        if (_commentStack.Count > 0)
            throw new InvalidOperationException("Unclosed CommentRange(s) remain.");
    }
}

class Program
{
    static void Main()
    {
        // Load the document (adjust the path as needed)
        Document doc = new Document("input.docx");

        // Ensure the document has the minimal required nodes
        doc.EnsureMinimum();

        // Validate that every opening tag has a matching closing tag
        var validator = new TagValidatorVisitor();
        doc.Accept(validator);
        validator.Verify();

        // Render the validated document (e.g., to PDF)
        doc.Save("output.pdf", SaveFormat.Pdf);
    }
}
