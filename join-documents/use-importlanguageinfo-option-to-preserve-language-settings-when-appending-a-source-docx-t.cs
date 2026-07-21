using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folder for all artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // Create destination document with French language setting.
        // -----------------------------------------------------------------
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Font.LocaleId = new CultureInfo("fr-FR").LCID; // Set language via LocaleId
        dstBuilder.Writeln("Destination text.");

        string dstPath = Path.Combine(artifactsDir, "Destination.docx");
        dstDoc.Save(dstPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create source document with English language setting.
        // -----------------------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Font.LocaleId = new CultureInfo("en-US").LCID; // Set language via LocaleId
        srcBuilder.Writeln("Source text.");

        string srcPath = Path.Combine(artifactsDir, "Source.docx");
        srcDoc.Save(srcPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the documents (demonstrates load rule usage).
        // -----------------------------------------------------------------
        Document destination = new Document(dstPath);
        Document source = new Document(srcPath);

        // Configure import options (ImportLanguageInfo is not required in this version).
        ImportFormatOptions importOptions = new ImportFormatOptions();

        // Append source to destination while keeping source formatting and language.
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting, importOptions);

        // Save the merged document.
        string mergedPath = Path.Combine(artifactsDir, "Merged.docx");
        destination.Save(mergedPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Validation: ensure the merged file exists.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        // Validation: check language IDs of runs.
        Document mergedDoc = new Document(mergedPath);
        bool hasEnglish = false;
        bool hasFrench = false;

        foreach (Run run in mergedDoc.GetChildNodes(NodeType.Run, true))
        {
            string text = run.Text.Trim();
            if (text == "Source text." && run.Font.LocaleId == new CultureInfo("en-US").LCID)
                hasEnglish = true;
            if (text == "Destination text." && run.Font.LocaleId == new CultureInfo("fr-FR").LCID)
                hasFrench = true;
        }

        if (!hasEnglish || !hasFrench)
            throw new InvalidOperationException("Language information was not preserved correctly.");

        // Program completed successfully.
    }
}
