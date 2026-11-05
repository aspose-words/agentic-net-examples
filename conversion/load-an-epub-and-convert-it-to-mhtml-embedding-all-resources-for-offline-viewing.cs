using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample EPUB document created for conversion to MHTML.");
        // Save the document as EPUB (input file).
        const string epubPath = "input.epub";
        sourceDoc.Save(epubPath, SaveFormat.Epub);

        // Verify the EPUB file was created.
        if (!File.Exists(epubPath))
            throw new InvalidOperationException("The input EPUB file was not created.");

        // Step 2: Load the EPUB file.
        Document epubDoc = new Document(epubPath);

        // Step 3: Convert and save as MHTML with embedded resources.
        const string mhtmlPath = "output.mht";
        epubDoc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Verify the MHTML file was created.
        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("The output MHTML file was not created.");

        // Optional: Indicate success (no interactive wait).
        Console.WriteLine("EPUB successfully converted to MHTML.");
    }
}
