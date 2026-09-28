using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class TableComparisonExample
{
    public static void Main()
    {
        // Prepare file paths in the current directory.
        string baseDir = Directory.GetCurrentDirectory();
        string docPath = Path.Combine(baseDir, "original.doc");
        string docxPath = Path.Combine(baseDir, "revised.docx");
        string resultPath = Path.Combine(baseDir, "comparisonResult.docx");

        // Create the original DOC with a simple 2x2 table.
        Document original = new Document();
        DocumentBuilder builderOrig = new DocumentBuilder(original);
        builderOrig.StartTable();
        builderOrig.InsertCell();
        builderOrig.Writeln("A1");
        builderOrig.InsertCell();
        builderOrig.Writeln("B1");
        builderOrig.EndRow();
        builderOrig.InsertCell();
        builderOrig.Writeln("A2");
        builderOrig.InsertCell();
        builderOrig.Writeln("B2");
        builderOrig.EndRow();
        builderOrig.EndTable();
        original.Save(docPath, SaveFormat.Doc);

        // Create the revised DOCX with a 2x3 table (added a column).
        Document revised = new Document();
        DocumentBuilder builderRev = new DocumentBuilder(revised);
        builderRev.StartTable();
        builderRev.InsertCell();
        builderRev.Writeln("A1");
        builderRev.InsertCell();
        builderRev.Writeln("B1");
        builderRev.InsertCell();
        builderRev.Writeln("C1");
        builderRev.EndRow();
        builderRev.InsertCell();
        builderRev.Writeln("A2");
        builderRev.InsertCell();
        builderRev.Writeln("B2");
        builderRev.InsertCell();
        builderRev.Writeln("C2");
        builderRev.EndRow();
        builderRev.EndTable();
        revised.Save(docxPath, SaveFormat.Docx);

        // Compare the original DOC with the revised DOCX.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that revisions were detected.
        int revisionCount = original.Revisions.Count;
        if (revisionCount == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison, but none were found.");
        }

        // Optionally, output the types of revisions detected (e.g., table structure changes).
        foreach (Revision rev in original.Revisions)
        {
            // For demonstration, write revision type to console.
            Console.WriteLine($"Revision Type: {rev.RevisionType}");
        }

        // Save the document that now contains revision markup.
        original.Save(resultPath, SaveFormat.Docx);
    }
}
