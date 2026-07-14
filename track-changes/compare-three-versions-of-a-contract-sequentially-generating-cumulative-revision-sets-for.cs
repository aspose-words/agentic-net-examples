using System;
using System.IO;
using Aspose.Words;

public class ContractRevisionDemo
{
    public static void Main()
    {
        // Directory to store generated documents.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the three contract versions.
        string originalPath = Path.Combine(outputDir, "Contract_Original.docx");
        string version1Path = Path.Combine(outputDir, "Contract_Version1.docx");
        string version2Path = Path.Combine(outputDir, "Contract_Version2.docx");

        // -----------------------------------------------------------------
        // 1. Create the original contract document (no revisions).
        // -----------------------------------------------------------------
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("Contract Title");
        builder.Writeln("Clause 1: The buyer shall pay $1000.");
        builder.Writeln("Clause 2: Delivery by June.");
        builder.Writeln("Clause 3: Warranty for 1 year.");
        original.Save(originalPath);

        // -----------------------------------------------------------------
        // 2. Create Version 1 with some tracked changes, then accept them
        //    so the saved file contains no pending revisions.
        // -----------------------------------------------------------------
        Document version1 = (Document)original.Clone(true);
        version1.StartTrackRevisions("Alice", DateTime.Now);
        DocumentBuilder builderV1 = new DocumentBuilder(version1);
        // Modify Clause 1 amount.
        builderV1.MoveTo(version1.FirstSection.Body.Paragraphs[1].GetChildNodes(NodeType.Run, true)[0]);
        builderV1.Write("The buyer shall pay $1200.");
        // Add a new clause.
        builderV1.Writeln("Clause 4: Confidentiality must be maintained.");
        version1.StopTrackRevisions();
        // Accept all revisions to produce a clean document for comparison.
        version1.AcceptAllRevisions();
        version1.Save(version1Path);

        // -----------------------------------------------------------------
        // 3. Create Version 2 with additional tracked changes, then accept them.
        // -----------------------------------------------------------------
        Document version2 = (Document)original.Clone(true);
        version2.StartTrackRevisions("Bob", DateTime.Now);
        DocumentBuilder builderV2 = new DocumentBuilder(version2);
        // Add a new clause.
        builderV2.Writeln("Clause 5: Disputes shall be resolved by arbitration.");
        // Extend Clause 2.
        builderV2.MoveTo(version2.FirstSection.Body.Paragraphs[2].GetChildNodes(NodeType.Run, true)[0]);
        builderV2.Write("Delivery by June. Late delivery incurs a penalty.");
        version2.StopTrackRevisions();
        // Accept all revisions to produce a clean document for comparison.
        version2.AcceptAllRevisions();
        version2.Save(version2Path);

        // -----------------------------------------------------------------
        // 4. First comparison: Original vs Version 1.
        //    The original document will receive revisions describing the changes.
        // -----------------------------------------------------------------
        Document docOriginal = new Document(originalPath);
        Document docV1 = new Document(version1Path);

        // Ensure both documents are revision‑free before comparison.
        if (docOriginal.HasRevisions || docV1.HasRevisions)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        docOriginal.Compare(docV1, "Comparer", DateTime.Now);
        string compare1Path = Path.Combine(outputDir, "Original_vs_Version1.docx");
        docOriginal.Save(compare1Path);
        Console.WriteLine($"Comparison 1 completed. Revisions generated: {docOriginal.Revisions.Count}");

        // -----------------------------------------------------------------
        // 5. Accept revisions to turn Original into Version 1.
        // -----------------------------------------------------------------
        docOriginal.AcceptAllRevisions();

        // -----------------------------------------------------------------
        // 6. Second comparison: Updated Original (now Version 1) vs Version 2.
        // -----------------------------------------------------------------
        Document docV2 = new Document(version2Path);
        // Both documents are revision‑free at this point.
        docOriginal.Compare(docV2, "Comparer2", DateTime.Now);
        string compare2Path = Path.Combine(outputDir, "Version1_vs_Version2.docx");
        docOriginal.Save(compare2Path);
        Console.WriteLine($"Comparison 2 completed. Additional revisions generated: {docOriginal.Revisions.Count}");

        // -----------------------------------------------------------------
        // 7. Accept all revisions to obtain the final contract.
        // -----------------------------------------------------------------
        docOriginal.AcceptAllRevisions();
        string finalPath = Path.Combine(outputDir, "Contract_Final.docx");
        docOriginal.Save(finalPath);
        Console.WriteLine($"Final contract saved with all revisions accepted: {finalPath}");
    }
}
