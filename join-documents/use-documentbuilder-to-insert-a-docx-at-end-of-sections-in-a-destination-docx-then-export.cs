using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // ---------- Create destination document ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Section 1
        destBuilder.Writeln("Destination Document - Section 1");
        destBuilder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        destBuilder.Writeln("Destination Document - Section 2");
        destBuilder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        destBuilder.Writeln("Destination Document - Section 3");

        string destPath = Path.Combine(outputDir, "Destination.docx");
        destDoc.Save(destPath, SaveFormat.Docx);

        // ---------- Create source document ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("=== Inserted Content Start ===");
        srcBuilder.Writeln("This is the content of the source DOCX.");
        srcBuilder.Writeln("=== Inserted Content End ===");

        string srcPath = Path.Combine(outputDir, "Source.docx");
        srcDoc.Save(srcPath, SaveFormat.Docx);

        // ---------- Insert source document at the end of each section ----------
        // Reload documents to simulate a real‑world scenario.
        Document destination = new Document(destPath);
        Document source = new Document(srcPath);
        DocumentBuilder builder = new DocumentBuilder(destination);

        // Preserve the original section count because inserting modifies the collection.
        int originalSectionCount = destination.Sections.Count;
        for (int i = 0; i < originalSectionCount; i++)
        {
            Section currentSection = destination.Sections[i];
            Paragraph lastParagraph = currentSection.Body.LastParagraph;

            // Move the cursor to the end of the current section.
            builder.MoveTo(lastParagraph);
            // Insert the source document while keeping its formatting.
            builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the merged result.
        string mergedPath = Path.Combine(outputDir, "Merged.docx");
        destination.Save(mergedPath, SaveFormat.Docx);

        // ---------- Validation ----------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        Document mergedDoc = new Document(mergedPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("This is the content of the source DOCX."))
            throw new InvalidOperationException("Merged document does not contain expected source content.");

        Console.WriteLine($"Merged document created at: {mergedPath}");
    }
}
