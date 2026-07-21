using System;
using System.IO;
using Aspose.Words;

namespace RevisionComparisonExample
{
    public class Program
    {
        public static void Main()
        {
            // Directory for generated files
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Paths for the three contract versions
            string v1Path = Path.Combine(outputDir, "ContractV1.docx");
            string v2Path = Path.Combine(outputDir, "ContractV2.docx");
            string v3Path = Path.Combine(outputDir, "ContractV3.docx");

            // -----------------------------------------------------------------
            // Create Version 1 of the contract
            // -----------------------------------------------------------------
            Document docV1 = new Document();
            DocumentBuilder builder = new DocumentBuilder(docV1);
            builder.Writeln("Contract");
            builder.Writeln("Party A: Alice");
            builder.Writeln("Party B: Bob");
            builder.Writeln("Amount: $1,000");
            builder.Writeln("Effective Date: 2023-01-01");
            docV1.Save(v1Path);

            // -----------------------------------------------------------------
            // Create Version 2 (modify amount and add a clause)
            // -----------------------------------------------------------------
            Document docV2 = new Document(v1Path);
            builder = new DocumentBuilder(docV2);
            // Change amount (add a new line for simplicity)
            builder.MoveToDocumentEnd();
            builder.Writeln("Amount: $1,500");
            // Add new clause
            builder.Writeln("Clause 1: Delivery shall be within 30 days.");
            docV2.Save(v2Path);

            // -----------------------------------------------------------------
            // Create Version 3 (change Party B and add another clause)
            // -----------------------------------------------------------------
            Document docV3 = new Document(v2Path);
            builder = new DocumentBuilder(docV3);
            // Change Party B (add a new line at the start)
            builder.MoveToDocumentStart();
            builder.Writeln("Party B: Charlie");
            // Add another clause
            builder.Writeln("Clause 2: Late delivery incurs a penalty of $100 per day.");
            docV3.Save(v3Path);

            // -----------------------------------------------------------------
            // First comparison: V1 vs V2
            // -----------------------------------------------------------------
            Document original = new Document(v1Path);
            Document editedV2 = new Document(v2Path);

            // Ensure no revisions exist before comparison
            if (original.Revisions.Count != 0 || editedV2.Revisions.Count != 0)
                throw new InvalidOperationException("Documents must not contain revisions before comparison.");

            // Compare to generate revisions
            original.Compare(editedV2, "Reviewer", DateTime.Now);
            string revV1V2Path = Path.Combine(outputDir, "Contract_V1_vs_V2.docx");
            original.Save(revV1V2Path);
            Console.WriteLine($"V1 vs V2 revisions: {original.Revisions.Count}");

            // Accept revisions to transform original into V2
            original.AcceptAllRevisions();

            // -----------------------------------------------------------------
            // Second comparison: (now V2) vs V3
            // -----------------------------------------------------------------
            Document editedV3 = new Document(v3Path);

            // Ensure the document we are comparing from has no pending revisions
            if (original.Revisions.Count != 0 || editedV3.Revisions.Count != 0)
                throw new InvalidOperationException("Documents must not contain revisions before second comparison.");

            // Compare to generate revisions for V2 -> V3
            original.Compare(editedV3, "Reviewer", DateTime.Now);
            string revV2V3Path = Path.Combine(outputDir, "Contract_V2_vs_V3.docx");
            original.Save(revV2V3Path);
            Console.WriteLine($"V2 vs V3 revisions: {original.Revisions.Count}");

            // Final acceptance (optional)
            original.AcceptAllRevisions();
            string finalPath = Path.Combine(outputDir, "Contract_Final.docx");
            original.Save(finalPath);
            Console.WriteLine("Final contract saved.");
        }
    }
}
