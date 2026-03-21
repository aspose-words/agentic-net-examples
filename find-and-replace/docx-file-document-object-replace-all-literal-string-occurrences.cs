using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

class ReplaceTextInDocx
{
    static void Main()
    {
        // Create a temporary directory for the demo files.
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(tempDir);

        // Path where the original and modified documents will be saved.
        string inputPath = Path.Combine(tempDir, "SourceDocument.docx");
        string outputPath = Path.Combine(tempDir, "ModifiedDocument.docx");

        // Create a new document with placeholder text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a sample document.");
        builder.Writeln("Please replace the placeholder: _FullName_");
        doc.Save(inputPath);

        // Reload the document from the file system (simulating a real‑world scenario).
        Document loadedDoc = new Document(inputPath);

        // The literal text to find and its replacement.
        string textToFind = "_FullName_";
        string replacementText = "John Doe";

        // Perform a case‑insensitive, whole‑document replace.
        loadedDoc.Range.Replace(textToFind, replacementText, new FindReplaceOptions(FindReplaceDirection.Forward));

        // Save the updated document.
        loadedDoc.Save(outputPath);

        Console.WriteLine($"Original document: {inputPath}");
        Console.WriteLine($"Modified document: {outputPath}");
    }
}
