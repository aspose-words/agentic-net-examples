using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names in the current working directory.
        string destPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docx");
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.rtf");
        string combinedPath = Path.Combine(Directory.GetCurrentDirectory(), "Combined.docx");

        // -----------------------------------------------------------------
        // 1. Create the destination DOCX document and add some content.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the content of the destination DOCX document.");

        // Save the destination document as DOCX (optional, demonstrates the format).
        destDoc.Save(destPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create the source RTF document and add some content.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the content of the source RTF document.");

        // Save the source document as RTF.
        sourceDoc.Save(sourcePath, SaveFormat.Rtf);

        // Load the RTF document back to ensure it is read as RTF.
        Document loadedSource = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Append the RTF document to the DOCX document using UseDestinationStyles.
        // -----------------------------------------------------------------
        destDoc.AppendDocument(loadedSource, ImportFormatMode.UseDestinationStyles);

        // -----------------------------------------------------------------
        // 4. Save the combined document as DOCX.
        // -----------------------------------------------------------------
        destDoc.Save(combinedPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 5. Validation: check that the file exists and contains both texts.
        // -----------------------------------------------------------------
        if (!File.Exists(combinedPath))
            throw new InvalidOperationException("The combined document was not created.");

        Document combinedDoc = new Document(combinedPath);
        string combinedText = combinedDoc.GetText();

        if (!combinedText.Contains("This is the content of the destination DOCX document.") ||
            !combinedText.Contains("This is the content of the source RTF document."))
        {
            throw new InvalidOperationException("The combined document does not contain expected content.");
        }

        // The example completes without requiring any user interaction.
    }
}
