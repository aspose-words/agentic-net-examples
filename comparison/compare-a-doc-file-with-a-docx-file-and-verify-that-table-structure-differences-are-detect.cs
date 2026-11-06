using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class TableComparisonExample
{
    public static void Main()
    {
        // Create the source DOC (old format) with a 2x2 table.
        Document docOld = new Document();
        DocumentBuilder builderOld = new DocumentBuilder(docOld);
        Table tableOld = builderOld.StartTable();
        builderOld.InsertCell();
        builderOld.Writeln("Old R1C1");
        builderOld.InsertCell();
        builderOld.Writeln("Old R1C2");
        builderOld.EndRow();
        builderOld.InsertCell();
        builderOld.Writeln("Old R2C1");
        builderOld.InsertCell();
        builderOld.Writeln("Old R2C2");
        builderOld.EndRow();
        builderOld.EndTable();

        string oldPath = Path.Combine(Directory.GetCurrentDirectory(), "source.doc");
        docOld.Save(oldPath, SaveFormat.Doc);

        // Create the revised DOCX (new format) with a 3x2 table (extra row).
        Document docNew = new Document();
        DocumentBuilder builderNew = new DocumentBuilder(docNew);
        Table tableNew = builderNew.StartTable();
        builderNew.InsertCell();
        builderNew.Writeln("New R1C1");
        builderNew.InsertCell();
        builderNew.Writeln("New R1C2");
        builderNew.EndRow();
        builderNew.InsertCell();
        builderNew.Writeln("New R2C1");
        builderNew.InsertCell();
        builderNew.Writeln("New R2C2");
        builderNew.EndRow();
        builderNew.InsertCell();
        builderNew.Writeln("New R3C1");
        builderNew.InsertCell();
        builderNew.Writeln("New R3C2");
        builderNew.EndRow();
        builderNew.EndTable();

        string newPath = Path.Combine(Directory.GetCurrentDirectory(), "revised.docx");
        docNew.Save(newPath, SaveFormat.Docx);

        // Load the documents from the saved files.
        Document baseDoc = new Document(oldPath);
        Document revisedDoc = new Document(newPath);

        // Compare the documents.
        baseDoc.Compare(revisedDoc, "TableComparer", DateTime.Now);

        // Verify that at least one revision (table structure difference) was detected.
        int revisionCount = baseDoc.Revisions.Count;
        if (revisionCount == 0)
        {
            throw new InvalidOperationException("Expected revisions due to table structure differences, but none were found.");
        }

        // Save the comparison result.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "comparison-result.docx");
        baseDoc.Save(resultPath, SaveFormat.Docx);

        // Output the number of revisions detected.
        Console.WriteLine($"Revisions detected: {revisionCount}");
    }
}
