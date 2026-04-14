using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string baseDir = Directory.GetCurrentDirectory();
        string sourcePath1 = Path.Combine(baseDir, "Source1.odt");
        string sourcePath2 = Path.Combine(baseDir, "Source2.odt");
        string sourcePath3 = Path.Combine(baseDir, "Source3.odt");
        string outputPath = Path.Combine(baseDir, "MergedOutput.docx");

        // Create first sample ODT document.
        Document srcDoc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(srcDoc1);
        builder1.Writeln("Content of first ODT document.");
        srcDoc1.Save(sourcePath1, SaveFormat.Odt);

        // Create second sample ODT document.
        Document srcDoc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(srcDoc2);
        builder2.Writeln("Content of second ODT document.");
        srcDoc2.Save(sourcePath2, SaveFormat.Odt);

        // Create third sample ODT document.
        Document srcDoc3 = new Document();
        DocumentBuilder builder3 = new DocumentBuilder(srcDoc3);
        builder3.Writeln("Content of third ODT document.");
        srcDoc3.Save(sourcePath3, SaveFormat.Odt);

        // Load the ODT documents.
        Document loadDoc1 = new Document(sourcePath1);
        Document loadDoc2 = new Document(sourcePath2);
        Document loadDoc3 = new Document(sourcePath3);

        // Create the destination document.
        Document dstDoc = new Document();

        // Append the source documents while preserving their formatting.
        dstDoc.AppendDocument(loadDoc1, ImportFormatMode.KeepSourceFormatting);
        dstDoc.AppendDocument(loadDoc2, ImportFormatMode.KeepSourceFormatting);
        dstDoc.AppendDocument(loadDoc3, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as DOCX.
        dstDoc.Save(outputPath, SaveFormat.Docx);

        // Validation: ensure the output file exists.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Merged output file was not created.");

        // Validation: ensure the merged document contains content from all sources.
        Document mergedDoc = new Document(outputPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("Content of first ODT document.") ||
            !mergedText.Contains("Content of second ODT document.") ||
            !mergedText.Contains("Content of third ODT document."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content from all source files.");
        }

        // Indicate successful completion.
        Console.WriteLine("Documents merged successfully. Output file: " + outputPath);
    }
}
