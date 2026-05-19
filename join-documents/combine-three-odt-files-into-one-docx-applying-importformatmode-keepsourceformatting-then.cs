using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the three source ODT files.
        string sourcePath1 = Path.Combine(outputDir, "Source1.odt");
        string sourcePath2 = Path.Combine(outputDir, "Source2.odt");
        string sourcePath3 = Path.Combine(outputDir, "Source3.odt");

        // Create first ODT document.
        Document srcDoc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(srcDoc1);
        builder1.Writeln("First ODT document content.");
        srcDoc1.Save(sourcePath1, SaveFormat.Odt);

        // Create second ODT document.
        Document srcDoc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(srcDoc2);
        builder2.Writeln("Second ODT document content.");
        srcDoc2.Save(sourcePath2, SaveFormat.Odt);

        // Create third ODT document.
        Document srcDoc3 = new Document();
        DocumentBuilder builder3 = new DocumentBuilder(srcDoc3);
        builder3.Writeln("Third ODT document content.");
        srcDoc3.Save(sourcePath3, SaveFormat.Odt);

        // Load the source documents.
        Document source1 = new Document(sourcePath1);
        Document source2 = new Document(sourcePath2);
        Document source3 = new Document(sourcePath3);

        // Destination document that will hold the merged content.
        Document destination = new Document();

        // Append each source document while preserving its original formatting.
        destination.AppendDocument(source1, ImportFormatMode.KeepSourceFormatting);
        destination.AppendDocument(source2, ImportFormatMode.KeepSourceFormatting);
        destination.AppendDocument(source3, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as DOCX.
        string mergedPath = Path.Combine(outputDir, "Merged.docx");
        destination.Save(mergedPath, SaveFormat.Docx);

        // Validation: ensure the file exists and contains text from all sources.
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged DOCX file was not created.");

        Document mergedDoc = new Document(mergedPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("First ODT document content.") ||
            !mergedText.Contains("Second ODT document content.") ||
            !mergedText.Contains("Third ODT document content."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }
    }
}
