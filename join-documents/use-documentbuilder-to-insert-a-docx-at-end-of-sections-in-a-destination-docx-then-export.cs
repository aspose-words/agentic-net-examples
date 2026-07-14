using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string sourcePath = "Source.docx";
        string destinationPath = "Destination.docx";
        string mergedPath = "Merged.docx";

        // -----------------------------------------------------------------
        // Create a sample source DOCX.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Source document content.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create a sample destination DOCX with two sections.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination section 1.");
        destBuilder.InsertBreak(BreakType.SectionBreakNewPage);
        destBuilder.Writeln("Destination section 2.");
        destDoc.Save(destinationPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the documents (loading from the files ensures the example works
        // even if the objects were modified elsewhere).
        // -----------------------------------------------------------------
        Document src = new Document(sourcePath);
        Document dst = new Document(destinationPath);

        // -----------------------------------------------------------------
        // Insert the source document at the end of the destination document.
        // Use DocumentBuilder to position the cursor at the document end.
        // -----------------------------------------------------------------
        DocumentBuilder builder = new DocumentBuilder(dst);
        builder.MoveToDocumentEnd();
        // Optional page break before the inserted content.
        builder.InsertBreak(BreakType.PageBreak);
        builder.InsertDocument(src, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Save the merged result.
        // -----------------------------------------------------------------
        dst.Save(mergedPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Validation: ensure the file exists and contains content from both docs.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException($"Merged file was not created: {mergedPath}");

        string mergedText = dst.GetText();

        if (!mergedText.Contains("Source document content.") ||
            !mergedText.Contains("Destination section 1.") ||
            !mergedText.Contains("Destination section 2."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }

        // Indicate successful completion (no interactive output required).
        Console.WriteLine("Merge completed successfully.");
    }
}
