using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string destDocPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docx");
        string srcRtfPath = Path.Combine(Directory.GetCurrentDirectory(), "Source.rtf");
        string mergedDocPath = Path.Combine(Directory.GetCurrentDirectory(), "Merged.docx");

        // -----------------------------------------------------------------
        // 1. Create the destination DOCX document.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination DOCX document.");
        destDoc.Save(destDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create the source RTF document.
        // -----------------------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the source RTF document.");
        srcDoc.Save(srcRtfPath, SaveFormat.Rtf);

        // -----------------------------------------------------------------
        // 3. Load both documents from disk.
        // -----------------------------------------------------------------
        Document destination = new Document(destDocPath);
        Document source = new Document(srcRtfPath);

        // -----------------------------------------------------------------
        // 4. Append the RTF document to the DOCX using destination styles.
        // -----------------------------------------------------------------
        destination.AppendDocument(source, ImportFormatMode.UseDestinationStyles);

        // -----------------------------------------------------------------
        // 5. Save the combined document as DOCX.
        // -----------------------------------------------------------------
        destination.Save(mergedDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 6. Validate that the merged file exists and contains content from both sources.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedDocPath))
            throw new InvalidOperationException("Merged document was not created.");

        Document merged = new Document(mergedDocPath);
        string mergedText = merged.GetText();

        if (!mergedText.Contains("destination DOCX document") ||
            !mergedText.Contains("source RTF document"))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }

        // The program finishes without requiring any user interaction.
    }
}
