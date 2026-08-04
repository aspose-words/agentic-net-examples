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

        // Paths for the sample documents.
        string destPath = Path.Combine(outputDir, "Destination.docx");
        string srcPath = Path.Combine(outputDir, "Source.docx");
        string mergedPath = Path.Combine(outputDir, "Merged.docx");

        // ---------- Create destination document with multiple sections ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        destBuilder.Writeln("Destination Section 1");
        destBuilder.InsertBreak(BreakType.SectionBreakNewPage);
        destBuilder.Writeln("Destination Section 2");
        destDoc.Save(destPath); // optional, just to have a physical file.

        // ---------- Create source document ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("Source Document Content");
        srcDoc.Save(srcPath); // optional.

        // ---------- Insert the source document at the end of each original section ----------
        int originalSectionCount = destDoc.Sections.Count; // capture before modifications.
        for (int i = 0; i < originalSectionCount; i++)
        {
            Section section = destDoc.Sections[i];

            // Move the builder to the last paragraph of the current section.
            destBuilder.MoveTo(section.Body.LastParagraph);

            // Optional page break before the inserted content.
            destBuilder.InsertBreak(BreakType.PageBreak);

            // Insert the source document preserving its formatting.
            destBuilder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // ---------- Save the merged document ----------
        destDoc.Save(mergedPath, SaveFormat.Docx);

        // ---------- Simple validation ----------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        Document mergedDoc = new Document(mergedPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("Destination Section 1") ||
            !mergedText.Contains("Destination Section 2") ||
            !mergedText.Contains("Source Document Content"))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }
    }
}
