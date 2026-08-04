using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for temporary files
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "JoinDocsExample");
        Directory.CreateDirectory(workDir);

        // Paths for the three source ODT files
        string sourcePath1 = Path.Combine(workDir, "Source1.odt");
        string sourcePath2 = Path.Combine(workDir, "Source2.odt");
        string sourcePath3 = Path.Combine(workDir, "Source3.odt");

        // Create first source ODT document
        Document srcDoc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(srcDoc1);
        builder1.Writeln("This is the content of the first ODT document.");
        srcDoc1.Save(sourcePath1, SaveFormat.Odt);

        // Create second source ODT document
        Document srcDoc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(srcDoc2);
        builder2.Writeln("Second ODT document comes here with different text.");
        srcDoc2.Save(sourcePath2, SaveFormat.Odt);

        // Create third source ODT document
        Document srcDoc3 = new Document();
        DocumentBuilder builder3 = new DocumentBuilder(srcDoc3);
        builder3.Writeln("Third ODT file adds its own paragraph.");
        srcDoc3.Save(sourcePath3, SaveFormat.Odt);

        // Load the ODT files (simulating real input files)
        Document loadDoc1 = new Document(sourcePath1);
        Document loadDoc2 = new Document(sourcePath2);
        Document loadDoc3 = new Document(sourcePath3);

        // Destination document that will hold the merged result
        Document dstDoc = new Document();

        // Append each source document preserving its original formatting
        dstDoc.AppendDocument(loadDoc1, ImportFormatMode.KeepSourceFormatting);
        dstDoc.AppendDocument(loadDoc2, ImportFormatMode.KeepSourceFormatting);
        dstDoc.AppendDocument(loadDoc3, ImportFormatMode.KeepSourceFormatting);

        // Path for the merged DOCX output
        string outputPath = Path.Combine(workDir, "MergedResult.docx");
        dstDoc.Save(outputPath, SaveFormat.Docx);

        // Validation: ensure the output file exists
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Merged document was not saved.");

        // Validation: ensure the merged document contains text from all sources
        string mergedText = dstDoc.GetText();
        if (!mergedText.Contains("first ODT document") ||
            !mergedText.Contains("Second ODT document") ||
            !mergedText.Contains("Third ODT file"))
        {
            throw new InvalidOperationException("Merged document does not contain expected content from all source files.");
        }

        // Clean up temporary files (optional)
        // File.Delete(sourcePath1);
        // File.Delete(sourcePath2);
        // File.Delete(sourcePath3);
        // File.Delete(outputPath);
        // Directory.Delete(workDir);
    }
}
