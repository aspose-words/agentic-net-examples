using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // ---------- Create the original contract ----------
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("Contract Agreement");
        builder.Writeln("This contract is between Party A and Party B.");
        builder.Writeln("Clause 1: The term is one year.");
        builder.Writeln("Clause 2: Payment shall be made monthly.");
        builder.Writeln("Clause 3: Confidentiality must be maintained.");

        // ---------- Create Version 1 (modify clause 2 and add clause 4) ----------
        Document version1 = (Document)original.Clone(true);
        // Modify Clause 2.
        Paragraph clause2V1 = version1.FirstSection.Body.Paragraphs[3]; // 0‑based index.
        clause2V1.Runs[0].Text = "Clause 2: Payment shall be made quarterly.";
        // Add Clause 4.
        DocumentBuilder b1 = new DocumentBuilder(version1);
        b1.MoveToDocumentEnd();
        b1.Writeln("Clause 4: Termination requires 30 days notice.");

        // ---------- Create Version 2 (modify clause 1, delete clause 3, add clause 5) ----------
        Document version2 = (Document)original.Clone(true);
        // Delete Clause 3.
        Paragraph clause3V2 = version2.FirstSection.Body.Paragraphs[4];
        clause3V2.Remove();
        // Modify Clause 1.
        Paragraph clause1V2 = version2.FirstSection.Body.Paragraphs[2];
        clause1V2.Runs[0].Text = "Clause 1: The term is two years.";
        // Add Clause 5.
        DocumentBuilder b2 = new DocumentBuilder(version2);
        b2.MoveToDocumentEnd();
        b2.Writeln("Clause 5: Governing law is XYZ.");

        // ---------- Comparison 1: Original vs Version 1 ----------
        Document compare1 = (Document)original.Clone(true);
        compare1.Compare(version1, "Reviewer1", DateTime.Now);
        string file1 = Path.Combine(outputDir, "Original_vs_Version1.docx");
        compare1.Save(file1);
        Console.WriteLine($"Revisions after Original vs Version1: {compare1.Revisions.Count}");

        // ---------- Comparison 2: Original vs Version 2 ----------
        Document compare2 = (Document)original.Clone(true);
        compare2.Compare(version2, "Reviewer2", DateTime.Now);
        string file2 = Path.Combine(outputDir, "Original_vs_Version2.docx");
        compare2.Save(file2);
        Console.WriteLine($"Revisions after Original vs Version2: {compare2.Revisions.Count}");

        // ---------- Comparison 3: Version 1 vs Version 2 ----------
        Document compare3 = (Document)version1.Clone(true);
        compare3.Compare(version2, "Reviewer3", DateTime.Now);
        string file3 = Path.Combine(outputDir, "Version1_vs_Version2.docx");
        compare3.Save(file3);
        Console.WriteLine($"Revisions after Version1 vs Version2: {compare3.Revisions.Count}");
    }
}
