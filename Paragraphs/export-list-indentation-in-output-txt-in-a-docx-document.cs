using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a three‑level numbered list.
        builder.ListFormat.ApplyNumberDefault(); // Level 0
        builder.Writeln("Item 1");
        builder.ListFormat.ListIndent();        // Level 1
        builder.Writeln("Item 2");
        builder.ListFormat.ListIndent();        // Level 2
        builder.Write("Item 3");

        // Configure TXT save options to indent each list level with three spaces.
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        txtOptions.ListIndentation.Character = ' ';
        txtOptions.ListIndentation.Count = 3;

        // Prepare output folder.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save as DOCX (optional) and as TXT with the configured indentation.
        string docxPath = Path.Combine(outputDir, "ListIndentation.docx");
        string txtPath = Path.Combine(outputDir, "ListIndentation.txt");
        doc.Save(docxPath);
        doc.Save(txtPath, txtOptions);

        // Display the resulting TXT content.
        Console.WriteLine("TXT output:");
        Console.WriteLine(File.ReadAllText(txtPath));
    }
}
