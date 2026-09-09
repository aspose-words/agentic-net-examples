using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define a folder for all temporary files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the source and destination documents.
        string destDocPath = Path.Combine(outputDir, "Destination.docx");
        string srcRtfPath = Path.Combine(outputDir, "Source.rtf");
        string combinedDocPath = Path.Combine(outputDir, "Combined.docx");

        // -----------------------------------------------------------------
        // Create the destination DOCX document.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination DOCX document.");
        destDoc.Save(destDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create the source RTF document.
        // -----------------------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the source RTF document.");
        srcDoc.Save(srcRtfPath, SaveFormat.Rtf);

        // -----------------------------------------------------------------
        // Load the documents from disk.
        // -----------------------------------------------------------------
        Document destination = new Document(destDocPath); // Loads DOCX.
        Document source = new Document(srcRtfPath);      // Loads RTF.

        // Append the RTF document to the DOCX using destination styles.
        destination.AppendDocument(source, ImportFormatMode.UseDestinationStyles);

        // Save the combined document as DOCX.
        destination.Save(combinedDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Validation: ensure the combined file exists and contains both texts.
        // -----------------------------------------------------------------
        if (!File.Exists(combinedDocPath))
            throw new InvalidOperationException("The combined document was not created.");

        Document combined = new Document(combinedDocPath);
        string combinedText = combined.GetText();

        if (!combinedText.Contains("This is the destination DOCX document.") ||
            !combinedText.Contains("This is the source RTF document."))
        {
            throw new InvalidOperationException("The combined document does not contain expected content.");
        }

        // Indicate successful completion.
        Console.WriteLine("Documents merged successfully. Output saved to: " + combinedDocPath);
    }
}
