using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string sourcePath = Path.Combine(outputDir, "Source.docx");
        string destinationPath = Path.Combine(outputDir, "Destination.docx");
        string mergedPath = Path.Combine(outputDir, "Merged.docx");

        // -----------------------------------------------------------------
        // 1. Create a source document (the DOCX to be inserted).
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("=== Source Document ===");
        srcBuilder.Writeln("This content comes from the source DOCX.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create a destination document with multiple sections.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Section 1
        destBuilder.Writeln("=== Destination Document - Section 1 ===");
        destBuilder.Writeln("First section content.");
        destBuilder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        destBuilder.Writeln("=== Destination Document - Section 2 ===");
        destBuilder.Writeln("Second section content.");
        // No extra break – we will insert the source after this section.

        destDoc.Save(destinationPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Load the documents (simulating a real‑world scenario where they
        //    already exist on disk) and insert the source at the end of the
        //    destination document using DocumentBuilder.
        // -----------------------------------------------------------------
        Document loadedDest = new Document(destinationPath);
        Document loadedSource = new Document(sourcePath);

        DocumentBuilder builder = new DocumentBuilder(loadedDest);
        // Move the cursor to the end of the document (after the last section).
        builder.MoveToDocumentEnd();
        // Optional: add a page break before the inserted content.
        builder.InsertBreak(BreakType.PageBreak);
        // Insert the source document preserving its original formatting.
        builder.InsertDocument(loadedSource, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 4. Save the merged result.
        // -----------------------------------------------------------------
        loadedDest.Save(mergedPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 5. Simple validation – ensure the merged file exists and contains
        //    text from both source and destination documents.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        Document verificationDoc = new Document(mergedPath);
        string mergedText = verificationDoc.GetText();

        if (!mergedText.Contains("Source Document") || !mergedText.Contains("Destination Document"))
            throw new InvalidOperationException("Merged document does not contain expected content.");

        // The program finishes without interactive prompts.
    }
}
