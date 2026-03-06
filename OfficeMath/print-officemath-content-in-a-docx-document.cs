using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the DOCX document that contains OfficeMath objects.
        Document doc = new Document("Office math.docx");

        // Traverse the document and collect the OfficeMath structure as plain text.
        OfficeMathPrinter printer = new OfficeMathPrinter();
        doc.Accept(printer);
        Console.WriteLine(printer.GetText());

        // The Document.Print() method is not available in the current Aspose.Words version.
        // If printing is required, use Aspose.Words.Printing.PrintDocument or another approach.
    }
}

// Visitor that extracts the structure and text of OfficeMath nodes.
class OfficeMathPrinter : DocumentVisitor
{
    private readonly StringBuilder _builder = new StringBuilder();
    private bool _insideOfficeMath;
    private int _depth;

    public string GetText() => _builder.ToString();

    public override VisitorAction VisitRun(Run run)
    {
        if (_insideOfficeMath)
            AppendLine($"[Run] \"{run.GetText()}\"");
        return VisitorAction.Continue;
    }

    public override VisitorAction VisitOfficeMathStart(OfficeMath officeMath)
    {
        AppendLine($"[OfficeMath start] Math object type: {officeMath.MathObjectType}");
        _depth++;
        _insideOfficeMath = true;
        return VisitorAction.Continue;
    }

    public override VisitorAction VisitOfficeMathEnd(OfficeMath officeMath)
    {
        _depth--;
        AppendLine("[OfficeMath end]");
        _insideOfficeMath = false;
        return VisitorAction.Continue;
    }

    private void AppendLine(string text)
    {
        for (int i = 0; i < _depth; i++) _builder.Append("|  ");
        _builder.AppendLine(text);
    }
}
