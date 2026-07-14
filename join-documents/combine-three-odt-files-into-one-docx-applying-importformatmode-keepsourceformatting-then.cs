using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string workFolder = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workFolder);

        // Paths for the three source ODT files.
        string odtPath1 = Path.Combine(workFolder, "Source1.odt");
        string odtPath2 = Path.Combine(workFolder, "Source2.odt");
        string odtPath3 = Path.Combine(workFolder, "Source3.odt");

        // Create first ODT document.
        Document srcDoc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(srcDoc1);
        builder1.Writeln("Content of the first ODT document.");
        srcDoc1.Save(odtPath1, SaveFormat.Odt);

        // Create second ODT document.
        Document srcDoc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(srcDoc2);
        builder2.Writeln("Content of the second ODT document.");
        srcDoc2.Save(odtPath2, SaveFormat.Odt);

        // Create third ODT document.
        Document srcDoc3 = new Document();
        DocumentBuilder builder3 = new DocumentBuilder(srcDoc3);
        builder3.Writeln("Content of the third ODT document.");
        srcDoc3.Save(odtPath3, SaveFormat.Odt);

        // Load the ODT documents back into Aspose.Words.
        Document loadedDoc1 = new Document(odtPath1);
        Document loadedDoc2 = new Document(odtPath2);
        Document loadedDoc3 = new Document(odtPath3);

        // Destination document that will hold the merged result.
        Document dstDoc = new Document();

        // Append each source document while preserving its original formatting.
        dstDoc.AppendDocument(loadedDoc1, ImportFormatMode.KeepSourceFormatting);
        dstDoc.AppendDocument(loadedDoc2, ImportFormatMode.KeepSourceFormatting);
        dstDoc.AppendDocument(loadedDoc3, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as DOCX.
        string outputPath = Path.Combine(workFolder, "Combined.docx");
        dstDoc.Save(outputPath, SaveFormat.Docx);

        // Validation: ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The merged DOCX file was not created.");

        // Validation: ensure the merged document contains text from all sources.
        Document resultDoc = new Document(outputPath);
        string resultText = resultDoc.GetText();

        if (!resultText.Contains("Content of the first ODT document.") ||
            !resultText.Contains("Content of the second ODT document.") ||
            !resultText.Contains("Content of the third ODT document."))
        {
            throw new InvalidOperationException("The merged document does not contain expected content from all source files.");
        }
    }
}
