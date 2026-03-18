using System;
using System.IO;
using System.Text;
using Aspose.Words;

class RevisionLogger
{
    private readonly StringBuilder _log = new StringBuilder();

    // Capture all current revisions in the document.
    public void Capture(Document doc)
    {
        foreach (Revision rev in doc.Revisions)
        {
            _log.AppendLine($"Author   : {rev.Author}");
            _log.AppendLine($"DateTime : {rev.DateTime}");
            _log.AppendLine($"Type     : {rev.RevisionType}");
            // ParentNode contains the node that was changed; GetText() returns its text.
            _log.AppendLine($"Text     : {rev.ParentNode?.GetText().Trim()}");
            _log.AppendLine(new string('-', 40));
        }
    }

    // Persist the collected log to a file.
    public void Save(string path)
    {
        File.WriteAllText(path, _log.ToString());
    }
}

class Program
{
    static void Main()
    {
        // ---------- Document creation ----------
        Document doc = new Document();                     // create a new document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Logger initialization ----------
        RevisionLogger logger = new RevisionLogger();

        // ---------- Start tracking revisions ----------
        // All subsequent changes will be recorded as revisions with the specified author and timestamp.
        doc.StartTrackRevisions("Auditor", DateTime.Now);

        // ---------- Perform edits ----------
        builder.Writeln("First line of the document.");
        builder.Writeln("Second line of the document.");
        builder.Writeln("Third line added later.");

        // ---------- Stop tracking ----------
        doc.StopTrackRevisions();

        // ---------- Capture revision metadata ----------
        logger.Capture(doc);

        // ---------- Save document ----------
        string docPath = "AuditedDocument.docx";
        doc.Save(docPath);                                 // save the edited document

        // ---------- Save revision log ----------
        string logPath = "RevisionLog.txt";
        logger.Save(logPath);                              // persist the audit log
    }
}
