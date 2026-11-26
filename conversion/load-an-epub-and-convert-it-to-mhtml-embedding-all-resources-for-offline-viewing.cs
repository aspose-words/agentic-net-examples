using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample EPUB document.");
        builder.Writeln("It will be converted to MHTML with embedded resources.");

        // Save the document as EPUB.
        string epubPath = "sample.epub";
        sourceDoc.Save(epubPath, SaveFormat.Epub);

        if (!File.Exists(epubPath) || new FileInfo(epubPath).Length == 0)
        {
            throw new InvalidOperationException("Failed to create the sample EPUB file.");
        }

        // Load the EPUB file.
        Document epubDoc = new Document(epubPath);

        // Convert and save as MHTML (embedded resources).
        string mhtmlPath = "output.mhtml";
        epubDoc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Validate the output.
        if (!File.Exists(mhtmlPath) || new FileInfo(mhtmlPath).Length == 0)
        {
            throw new InvalidOperationException("MHTML conversion failed; output file is missing or empty.");
        }

        // Optionally, clean up the intermediate EPUB file.
        // File.Delete(epubPath);
    }
}
