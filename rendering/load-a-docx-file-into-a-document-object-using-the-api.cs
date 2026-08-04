using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for the sample files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Path of the DOCX file that will be created and later loaded.
        string docPath = Path.Combine(dataDir, "Sample.docx");

        // -----------------------------------------------------------------
        // Create a simple DOCX document and save it to the file system.
        // -----------------------------------------------------------------
        Document doc = new Document();                     // Create a blank document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");           // Add some text.
        doc.Save(docPath);                                // Save; format inferred from ".docx".

        // -----------------------------------------------------------------
        // Load the previously saved DOCX file into a new Document object.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);       // Load using the file name.

        // Simple verification: output the text contained in the loaded document.
        string loadedText = loadedDoc.GetText().Trim();
        Console.WriteLine($"Loaded document text: {loadedText}");
    }
}
