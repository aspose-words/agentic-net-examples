using System;
using System.IO;
using Aspose.Words;

public class AppendRtfToDocx
{
    public static void Main()
    {
        // Define output directory and file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string destPath = Path.Combine(outputDir, "Destination.docx");
        string srcPath = Path.Combine(outputDir, "Source.rtf");
        string mergedPath = Path.Combine(outputDir, "Merged.docx");

        // Create the destination DOCX document.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination DOCX document.");
        destDoc.Save(destPath, SaveFormat.Docx);

        // Create the source RTF document.
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the source RTF document.");
        srcDoc.Save(srcPath, SaveFormat.Rtf);

        // Load the documents (format detection is automatic).
        Document destination = new Document(destPath);
        Document source = new Document(srcPath);

        // Append the RTF document to the DOCX using destination styles.
        destination.AppendDocument(source, ImportFormatMode.UseDestinationStyles);

        // Save the combined document as DOCX.
        destination.Save(mergedPath, SaveFormat.Docx);

        // Validation: ensure the merged file exists.
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        // Validation: ensure the merged document contains content from both sources.
        Document mergedDoc = new Document(mergedPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("This is the destination DOCX document.") ||
            !mergedText.Contains("This is the source RTF document."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }

        // Optionally, clean up intermediate files (comment out if inspection is needed).
        // File.Delete(destPath);
        // File.Delete(srcPath);
    }
}
