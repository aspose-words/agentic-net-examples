using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class TableComparisonDemo
{
    public static void Main()
    {
        // Prepare a folder for the sample files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "TableComparisonDemo");
        Directory.CreateDirectory(workDir);

        // Create the first document (DOC) with a 2x2 table.
        Document docOld = new Document();
        DocumentBuilder builderOld = new DocumentBuilder(docOld);
        builderOld.StartTable();
        builderOld.InsertCell();
        builderOld.Write("Old R1C1");
        builderOld.InsertCell();
        builderOld.Write("Old R1C2");
        builderOld.EndRow();
        builderOld.InsertCell();
        builderOld.Write("Old R2C1");
        builderOld.InsertCell();
        builderOld.Write("Old R2C2");
        builderOld.EndRow();
        builderOld.EndTable();
        string oldPath = Path.Combine(workDir, "OldDocument.doc");
        docOld.Save(oldPath, SaveFormat.Doc);

        // Create the second document (DOCX) with a 3x2 table (extra row).
        Document docNew = new Document();
        DocumentBuilder builderNew = new DocumentBuilder(docNew);
        builderNew.StartTable();
        builderNew.InsertCell();
        builderNew.Write("New R1C1");
        builderNew.InsertCell();
        builderNew.Write("New R1C2");
        builderNew.EndRow();
        builderNew.InsertCell();
        builderNew.Write("New R2C1");
        builderNew.InsertCell();
        builderNew.Write("New R2C2");
        builderNew.EndRow();
        // Extra row to create a structural difference.
        builderNew.InsertCell();
        builderNew.Write("New R3C1");
        builderNew.InsertCell();
        builderNew.Write("New R3C2");
        builderNew.EndRow();
        builderNew.EndTable();
        string newPath = Path.Combine(workDir, "NewDocument.docx");
        docNew.Save(newPath, SaveFormat.Docx);

        // Load the documents for comparison.
        Document sourceDoc = new Document(oldPath);
        Document targetDoc = new Document(newPath);

        // Perform comparison.
        sourceDoc.Compare(targetDoc, "TableComparer", DateTime.Now);

        // Verify that revisions were created due to table differences.
        int revisionCount = sourceDoc.Revisions.Count;
        if (revisionCount == 0)
        {
            throw new InvalidOperationException("Expected revisions due to table structure differences, but none were found.");
        }

        // Output the number of revisions detected.
        Console.WriteLine($"Revisions detected: {revisionCount}");

        // Save the compared document showing revisions.
        string resultPath = Path.Combine(workDir, "ComparedResult.docx");
        sourceDoc.Save(resultPath, SaveFormat.Docx);
    }
}
