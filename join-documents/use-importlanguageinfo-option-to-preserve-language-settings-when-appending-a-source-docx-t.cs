using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Set up output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths for the sample documents.
        string destPath = Path.Combine(outputDir, "Destination.docx");
        string srcPath = Path.Combine(outputDir, "Source.docx");
        string mergedPath = Path.Combine(outputDir, "Merged.docx");

        // ---------- Create destination document (French language) ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Font.LocaleId = CultureInfo.GetCultureInfo("fr-FR").LCID;
        destBuilder.Writeln("Texte du document de destination."); // French text.
        destDoc.Save(destPath, SaveFormat.Docx);

        // ---------- Create source document (English language) ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Font.LocaleId = CultureInfo.GetCultureInfo("en-US").LCID;
        srcBuilder.Writeln("Source document text."); // English text.
        srcDoc.Save(srcPath, SaveFormat.Docx);

        // Load the documents (they were just saved to disk).
        Document destination = new Document(destPath);
        Document source = new Document(srcPath);

        // Append the source document while keeping its formatting (including language settings).
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);
        destination.Save(mergedPath, SaveFormat.Docx);

        // ---------- Validation ----------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        string mergedText = new Document(mergedPath).GetText();
        if (!mergedText.Contains("Texte du document de destination.") ||
            !mergedText.Contains("Source document text."))
            throw new InvalidOperationException("Merged document does not contain expected content.");

        Console.WriteLine("Documents merged successfully. Output saved to:");
        Console.WriteLine(mergedPath);
    }
}
