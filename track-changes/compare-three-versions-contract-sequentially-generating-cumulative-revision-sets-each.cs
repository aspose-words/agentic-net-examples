using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

class ContractRevisionDemo
{
    static void Main()
    {
        // Create a temporary folder for the demo files.
        string tempDir = Path.Combine(Path.GetTempPath(), "ContractRevisionDemo");
        Directory.CreateDirectory(tempDir);

        // Paths to the three contract versions.
        string version1Path = Path.Combine(tempDir, "Contract_v1.docx");
        string version2Path = Path.Combine(tempDir, "Contract_v2.docx");
        string version3Path = Path.Combine(tempDir, "Contract_v3.docx");

        // Helper to create a simple document with given text.
        void CreateDoc(string path, string text)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln(text);
            doc.Save(path);
        }

        // Create three versions with slight differences.
        CreateDoc(version1Path, "Contract Version 1\nClause A: Original text.");
        CreateDoc(version2Path, "Contract Version 2\nClause A: Modified text.");
        CreateDoc(version3Path, "Contract Version 3\nClause A: Modified text.\nClause B: New clause added.");

        // -----------------------------------------------------------------
        // 1. Compare V1 with V2 and save the revision set.
        // -----------------------------------------------------------------
        var revDocV1V2 = new Document(version1Path);
        revDocV1V2.Compare(new Document(version2Path), "AuthorV1V2", DateTime.Now);
        revDocV1V2.Save(Path.Combine(tempDir, "Revision_V1_V2.docx"));

        // -----------------------------------------------------------------
        // 2. Compare V2 with V3 and save the revision set.
        // -----------------------------------------------------------------
        var revDocV2V3 = new Document(version2Path);
        revDocV2V3.Compare(new Document(version3Path), "AuthorV2V3", DateTime.Now);
        revDocV2V3.Save(Path.Combine(tempDir, "Revision_V2_V3.docx"));

        // -----------------------------------------------------------------
        // 3. Create a cumulative revision document:
        //    - Start with V1.
        //    - Compare V1 with V2, accept the revisions to turn V1 into V2.
        //    - Compare the updated document (now V2) with V3, leaving revisions unaccepted.
        // -----------------------------------------------------------------
        var cumulativeDoc = new Document(version1Path);
        cumulativeDoc.Compare(new Document(version2Path), "AuthorV1V2", DateTime.Now);
        cumulativeDoc.Revisions.AcceptAll(); // Accept first set of revisions.

        cumulativeDoc.Compare(new Document(version3Path), "AuthorV2V3", DateTime.Now);
        // Do NOT accept the second set of revisions – they remain as cumulative changes.
        cumulativeDoc.Save(Path.Combine(tempDir, "Revision_Cumulative.docx"));

        Console.WriteLine($"Demo files created in: {tempDir}");
    }
}
