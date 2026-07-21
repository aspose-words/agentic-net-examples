using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string baseDir = Directory.GetCurrentDirectory();
        string destPath = Path.Combine(baseDir, "Destination.docx");
        string srcPath = Path.Combine(baseDir, "Source.rtf");
        string combinedPath = Path.Combine(baseDir, "Combined.docx");

        // -----------------------------------------------------------------
        // Create the destination DOCX document.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination DOCX content.");
        destDoc.Save(destPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create the source RTF document.
        // -----------------------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("Source RTF content.");
        srcDoc.Save(srcPath, SaveFormat.Rtf);

        // -----------------------------------------------------------------
        // Load the documents from disk.
        // -----------------------------------------------------------------
        Document destination = new Document(destPath);
        Document source = new Document(srcPath);

        // -----------------------------------------------------------------
        // Append the RTF document to the DOCX using destination styles.
        // -----------------------------------------------------------------
        destination.AppendDocument(source, ImportFormatMode.UseDestinationStyles);

        // -----------------------------------------------------------------
        // Save the combined document as DOCX.
        // -----------------------------------------------------------------
        destination.Save(combinedPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Validation: ensure the combined file exists and contains both texts.
        // -----------------------------------------------------------------
        if (!File.Exists(combinedPath))
            throw new InvalidOperationException("Combined document was not created.");

        Document combined = new Document(combinedPath);
        string combinedText = combined.GetText();

        if (!combinedText.Contains("Destination DOCX content.") ||
            !combinedText.Contains("Source RTF content."))
        {
            throw new InvalidOperationException("Combined document does not contain expected content.");
        }

        // The program finishes without requiring user interaction.
    }
}
