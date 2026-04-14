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

        // Create the destination document with two sections.
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Writeln("Destination Document - Section 1");
        dstBuilder.InsertBreak(BreakType.SectionBreakNewPage);
        dstBuilder.Writeln("Destination Document - Section 2");

        // Create the source document that will be inserted.
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("Source Document Content");

        // Insert the source document at the end of the destination document (after all sections).
        dstBuilder.MoveToDocumentEnd();
        dstBuilder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the merged result.
        string mergedPath = Path.Combine(outputDir, "MergedResult.docx");
        dstDoc.Save(mergedPath, SaveFormat.Docx);

        // Validation: ensure the file was created.
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not saved.");

        // Validation: ensure the merged document contains text from both source and destination.
        Document mergedDoc = new Document(mergedPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("Destination Document - Section 1") ||
            !mergedText.Contains("Destination Document - Section 2") ||
            !mergedText.Contains("Source Document Content"))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }

        // Indicate successful completion (no interactive output required).
        Console.WriteLine("Document merging completed successfully.");
    }
}
