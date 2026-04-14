using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class PreserveLanguageOnAppend
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File names.
        string destPath = Path.Combine(outputDir, "Destination.docx");
        string srcPath = Path.Combine(outputDir, "Source.docx");
        string mergedPath = Path.Combine(outputDir, "Merged.docx");

        // Create destination document with English language setting.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Font.LocaleId = 1033; // English - United States
        destBuilder.Writeln("This is the destination document (English).");
        destDoc.Save(destPath);

        // Create source document with Japanese language setting.
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Font.LocaleId = 1041; // Japanese
        srcBuilder.Writeln("これはソース文書です (Japanese).");
        srcDoc.Save(srcPath);

        // Load the documents back (simulating real‑world scenario).
        Document destination = new Document(destPath);
        Document source = new Document(srcPath);

        // Set ImportFormatOptions – no special language flag is required;
        // language information is part of the formatting that is preserved
        // when using KeepSourceFormatting.
        ImportFormatOptions importOptions = new ImportFormatOptions();

        // Append source to destination while keeping source formatting (including language).
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting, importOptions);
        destination.Save(mergedPath);

        // Validation: ensure the merged file exists.
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        // Validation: ensure both pieces of text are present.
        Document merged = new Document(mergedPath);
        string mergedText = merged.GetText();

        if (!mergedText.Contains("destination document"))
            throw new InvalidOperationException("Destination text missing in merged document.");

        if (!mergedText.Contains("ソース文書です"))
            throw new InvalidOperationException("Source text missing in merged document.");

        // Confirmation.
        Console.WriteLine("Documents merged successfully. Output saved to:");
        Console.WriteLine(mergedPath);
    }
}
