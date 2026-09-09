using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create three sample ODT source documents.
        string doc1Path = Path.Combine(outputDir, "Doc1.odt");
        string doc2Path = Path.Combine(outputDir, "Doc2.odt");
        string doc3Path = Path.Combine(outputDir, "Doc3.odt");

        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("This is the first ODT document.");
        doc1.Save(doc1Path, SaveFormat.Odt);

        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("This is the second ODT document.");
        doc2.Save(doc2Path, SaveFormat.Odt);

        Document doc3 = new Document();
        DocumentBuilder builder3 = new DocumentBuilder(doc3);
        builder3.Writeln("This is the third ODT document.");
        doc3.Save(doc3Path, SaveFormat.Odt);

        // Load the ODT documents.
        Document src1 = new Document(doc1Path);
        Document src2 = new Document(doc2Path);
        Document src3 = new Document(doc3Path);

        // Create the destination document and append the sources preserving their formatting.
        Document dst = new Document();
        dst.AppendDocument(src1, ImportFormatMode.KeepSourceFormatting);
        dst.AppendDocument(src2, ImportFormatMode.KeepSourceFormatting);
        dst.AppendDocument(src3, ImportFormatMode.KeepSourceFormatting);

        // Save the merged result as DOCX.
        string mergedPath = Path.Combine(outputDir, "Merged.docx");
        dst.Save(mergedPath, SaveFormat.Docx);

        // Validation: ensure the file exists and contains content from all sources.
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        Document merged = new Document(mergedPath);
        string mergedText = merged.GetText();

        if (!mergedText.Contains("first ODT") ||
            !mergedText.Contains("second ODT") ||
            !mergedText.Contains("third ODT"))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }

        Console.WriteLine("Documents merged successfully. Output file: " + mergedPath);
    }
}
