using System;
using System.IO;
using Aspose.Words;

public class ContractRevisionDemo
{
    public static void Main()
    {
        // Directory to store temporary documents.
        string outputDir = Directory.GetCurrentDirectory();

        // Paths for the three contract versions.
        string v1Path = Path.Combine(outputDir, "Contract_V1.docx");
        string v2Path = Path.Combine(outputDir, "Contract_V2.docx");
        string v3Path = Path.Combine(outputDir, "Contract_V3.docx");

        // -------------------------
        // Create Version 1 of contract
        // -------------------------
        Document docV1 = new Document();
        DocumentBuilder builder = new DocumentBuilder(docV1);
        builder.Writeln("Contract Agreement");
        builder.Writeln("Party A: Alice");
        builder.Writeln("Party B: Bob");
        builder.Writeln("Effective Date: 2023-01-01");
        builder.Writeln("Terms: The service will be provided for 12 months.");
        docV1.Save(v1Path);

        // -------------------------
        // Create Version 2 of contract (some changes)
        // -------------------------
        Document docV2 = new Document();
        builder = new DocumentBuilder(docV2);
        builder.Writeln("Contract Agreement");
        builder.Writeln("Party A: Alice");
        builder.Writeln("Party B: Bob");
        builder.Writeln("Effective Date: 2023-02-01"); // date changed
        builder.Writeln("Terms: The service will be provided for 24 months."); // term changed
        docV2.Save(v2Path);

        // -------------------------
        // Create Version 3 of contract (further changes)
        // -------------------------
        Document docV3 = new Document();
        builder = new DocumentBuilder(docV3);
        builder.Writeln("Contract Agreement");
        builder.Writeln("Party A: Alice");
        builder.Writeln("Party B: Charlie"); // party B changed
        builder.Writeln("Effective Date: 2023-02-01");
        builder.Writeln("Terms: The service will be provided for 24 months.");
        builder.Writeln("Additional Clause: Confidentiality must be maintained.");
        docV3.Save(v3Path);

        // -------------------------------------------------
        // First comparison: Version 1 vs Version 2
        // -------------------------------------------------
        Document compare1 = new Document(v1Path);
        Document docV2ForCompare = new Document(v2Path);

        // Ensure both documents have no revisions before comparison (they don't).
        compare1.Compare(docV2ForCompare, "Reviewer1", DateTime.Now);
        Console.WriteLine($"Comparison 1 revisions count: {compare1.Revisions.Count}");

        string comp1Path = Path.Combine(outputDir, "Contract_Comparison_V1_vs_V2.docx");
        compare1.Save(comp1Path);

        // -------------------------------------------------
        // Second comparison: Version 2 vs Version 3
        // -------------------------------------------------
        Document compare2 = new Document(v2Path);
        Document docV3ForCompare = new Document(v3Path);

        compare2.Compare(docV3ForCompare, "Reviewer2", DateTime.Now);
        Console.WriteLine($"Comparison 2 revisions count: {compare2.Revisions.Count}");

        string comp2Path = Path.Combine(outputDir, "Contract_Comparison_V2_vs_V3.docx");
        compare2.Save(comp2Path);
    }
}
