using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define a folder for temporary files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Path of the sample DOCX file.
        string samplePath = Path.Combine(artifactsDir, "Sample.docx");

        // 1. Create a simple DOCX document.
        Document createdDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(createdDoc);
        builder.Writeln("Hello Aspose.Words!");
        createdDoc.Save(samplePath); // Persist the document so it can be loaded later.

        // 2. Load the DOCX file into a new Document object.
        Document loadedDoc = new Document(samplePath);

        // 3. Output the loaded document's text to verify successful loading.
        Console.WriteLine("Loaded document text:");
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}
