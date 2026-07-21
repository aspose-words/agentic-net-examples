using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folder for all files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "JoinDocumentsExample");
        Directory.CreateDirectory(outputDir);

        // Paths for the three source ODT files.
        string sourcePath1 = Path.Combine(outputDir, "Source1.odt");
        string sourcePath2 = Path.Combine(outputDir, "Source2.odt");
        string sourcePath3 = Path.Combine(outputDir, "Source3.odt");

        // Path for the merged DOCX file.
        string mergedPath = Path.Combine(outputDir, "MergedDocument.docx");

        // Create first ODT source document.
        Document srcDoc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(srcDoc1);
        builder1.Writeln("First ODT document content.");
        srcDoc1.Save(sourcePath1, SaveFormat.Odt);

        // Create second ODT source document.
        Document srcDoc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(srcDoc2);
        builder2.Writeln("Second ODT document content.");
        srcDoc2.Save(sourcePath2, SaveFormat.Odt);

        // Create third ODT source document.
        Document srcDoc3 = new Document();
        DocumentBuilder builder3 = new DocumentBuilder(srcDoc3);
        builder3.Writeln("Third ODT document content.");
        srcDoc3.Save(sourcePath3, SaveFormat.Odt);

        // Load the source documents (could reuse the in‑memory objects, but loading demonstrates the workflow).
        Document source1 = new Document(sourcePath1);
        Document source2 = new Document(sourcePath2);
        Document source3 = new Document(sourcePath3);

        // Destination document – starts as a blank document.
        Document destination = new Document();

        // Append each source document while preserving its original formatting.
        destination.AppendDocument(source1, ImportFormatMode.KeepSourceFormatting);
        destination.AppendDocument(source2, ImportFormatMode.KeepSourceFormatting);
        destination.AppendDocument(source3, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as DOCX.
        destination.Save(mergedPath, SaveFormat.Docx);

        // Validation: ensure the merged file exists.
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("The merged DOCX file was not created.");

        // Validation: ensure the merged document contains text from all three source documents.
        Document mergedDoc = new Document(mergedPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("First ODT document content.") ||
            !mergedText.Contains("Second ODT document content.") ||
            !mergedText.Contains("Third ODT document content."))
        {
            throw new InvalidOperationException("The merged document does not contain all expected content.");
        }

        // Indicate successful completion.
        Console.WriteLine("Documents merged successfully. Output located at:");
        Console.WriteLine(mergedPath);
    }
}
