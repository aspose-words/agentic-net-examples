using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths
        string destPath = Path.Combine(outputDir, "Destination.docx");
        string srcPath = Path.Combine(outputDir, "Source.docx");
        string mergedPath = Path.Combine(outputDir, "Merged.docx");

        // ---------- Create destination document ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        // Set French language using LocaleId (LCID)
        destBuilder.Font.LocaleId = new CultureInfo("fr-FR").LCID;
        destBuilder.Writeln("Ceci est le texte du document de destination.");
        destDoc.Save(destPath, SaveFormat.Docx);

        // ---------- Create source document ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        // Set English language using LocaleId (LCID)
        srcBuilder.Font.LocaleId = new CultureInfo("en-US").LCID;
        srcBuilder.Writeln("This is the source document text.");
        srcDoc.Save(srcPath, SaveFormat.Docx);

        // ---------- Load documents ----------
        Document destination = new Document(destPath);
        Document source = new Document(srcPath);

        // ---------- Append with language preservation ----------
        // ImportFormatOptions does not have an ImportLanguageInfo property.
        // Language information (LocaleId) is preserved automatically when appending.
        ImportFormatOptions importOptions = new ImportFormatOptions();

        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting, importOptions);
        destination.Save(mergedPath, SaveFormat.Docx);

        // ---------- Validation ----------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        // Ensure the merged document has at least two sections (original + appended)
        if (destination.Sections.Count < 2)
            throw new InvalidOperationException("Appended document does not contain a second section.");

        // The appended text resides in the second section (index 1)
        Run appendedRun = (Run)destination.Sections[1].Body.FirstParagraph.Runs[0];
        // Convert LocaleId back to culture name for display
        string appendedCulture = new CultureInfo(appendedRun.Font.LocaleId).Name;
        Console.WriteLine($"Appended run language: {appendedCulture}");
    }
}
