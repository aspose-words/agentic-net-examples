using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Folder for output files (relative to the executable directory).
        const string outputDir = "Output";
        System.IO.Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create three versions of a contract: V1, V2 and V3.
        // -----------------------------------------------------------------
        Document docV1 = new Document();
        DocumentBuilder builder = new DocumentBuilder(docV1);
        builder.Writeln("Contract Version 1");
        builder.Writeln("Clause A: The buyer shall pay $10,000.");
        builder.Writeln("Clause B: Delivery shall occur within 30 days.");
        docV1.Save(System.IO.Path.Combine(outputDir, "ContractV1.docx"));

        Document docV2 = new Document();
        builder = new DocumentBuilder(docV2);
        builder.Writeln("Contract Version 2");
        // Modified Clause A.
        builder.Writeln("Clause A: The buyer shall pay $12,000.");
        builder.Writeln("Clause B: Delivery shall occur within 30 days.");
        docV2.Save(System.IO.Path.Combine(outputDir, "ContractV2.docx"));

        Document docV3 = new Document();
        builder = new DocumentBuilder(docV3);
        builder.Writeln("Contract Version 3");
        builder.Writeln("Clause A: The buyer shall pay $12,000.");
        // Modified Clause B.
        builder.Writeln("Clause B: Delivery shall occur within 45 days.");
        docV3.Save(System.IO.Path.Combine(outputDir, "ContractV3.docx"));

        // -----------------------------------------------------------------
        // 2. Compare V1 with V2 – generate revisions in V1.
        // -----------------------------------------------------------------
        // Ensure both documents have no revisions before comparison.
        if (docV1.Revisions.Count != 0 || docV2.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must be revision‑free before comparison.");

        docV1.Compare(docV2, "Reviewer_V2", DateTime.Now);
        string compare12Path = System.IO.Path.Combine(outputDir, "Contract_V1_vs_V2.docx");
        docV1.Save(compare12Path);

        Console.WriteLine("Comparison V1 vs V2:");
        Console.WriteLine($"  Revisions count: {docV1.Revisions.Count}");
        foreach (Revision rev in docV1.Revisions)
        {
            Console.WriteLine($"  - Type: {rev.RevisionType}, Author: {rev.Author}, Text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Accept all revisions so that V1 becomes V2.
        docV1.AcceptAllRevisions();

        // -----------------------------------------------------------------
        // 3. Compare (now updated) V1 (which equals V2) with V3 – generate revisions.
        // -----------------------------------------------------------------
        // At this point docV1 has the content of V2 and no revisions.
        if (docV1.Revisions.Count != 0 || docV3.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must be revision‑free before second comparison.");

        docV1.Compare(docV3, "Reviewer_V3", DateTime.Now);
        string compare23Path = System.IO.Path.Combine(outputDir, "Contract_V2_vs_V3.docx");
        docV1.Save(compare23Path);

        Console.WriteLine("Comparison V2 vs V3:");
        Console.WriteLine($"  Revisions count: {docV1.Revisions.Count}");
        foreach (Revision rev in docV1.Revisions)
        {
            Console.WriteLine($"  - Type: {rev.RevisionType}, Author: {rev.Author}, Text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // -----------------------------------------------------------------
        // 4. (Optional) Accept all revisions to obtain the final document (V3).
        // -----------------------------------------------------------------
        docV1.AcceptAllRevisions();
        string finalPath = System.IO.Path.Combine(outputDir, "Contract_Final_V3.docx");
        docV1.Save(finalPath);

        // End of example – all files are written to the Output folder.
    }
}
