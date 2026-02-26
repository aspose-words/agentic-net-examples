using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportListIndentation
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a three‑level numbered list.
        builder.ListFormat.ApplyNumberDefault();   // Level 0
        builder.Writeln("Item 1");                // 1. Item 1
        builder.ListFormat.ListIndent();          // Level 1
        builder.Writeln("Item 2");                // a. Item 2
        builder.ListFormat.ListIndent();          // Level 2
        builder.Write("Item 3");                  // i. Item 3

        // Configure TXT save options to indent list levels.
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        txtOptions.ListIndentation.Character = ' '; // Use space for indentation.
        txtOptions.ListIndentation.Count = 3;       // Three spaces per level.

        // Define output paths.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string txtPath = Path.Combine(artifactsDir, "ListIndentation.txt");
        string docxPath = Path.Combine(artifactsDir, "ListIndentation.docx");

        // Save the document as DOCX (optional, to keep the source file).
        doc.Save(docxPath);

        // Save the same document as plain text using the configured indentation.
        doc.Save(txtPath, txtOptions);

        // Output the resulting text to the console for verification.
        Console.WriteLine(File.ReadAllText(txtPath));
    }
}
