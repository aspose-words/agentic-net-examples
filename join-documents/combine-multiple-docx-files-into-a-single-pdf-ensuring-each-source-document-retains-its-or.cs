using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string doc1Path = Path.Combine(Directory.GetCurrentDirectory(), "Sample1.docx");
        string doc2Path = Path.Combine(Directory.GetCurrentDirectory(), "Sample2.docx");
        string mergedPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.pdf");

        // Create first sample DOCX.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("This is the first sample document.");
        builder1.Writeln("It uses the default style.");
        doc1.Save(doc1Path, SaveFormat.Docx);

        // Create second sample DOCX with a different style.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Font.Name = "Courier New";
        builder2.Font.Size = 14;
        builder2.Font.Color = System.Drawing.Color.DarkBlue;
        builder2.Writeln("This is the second sample document.");
        builder2.Writeln("It uses a custom font and color.");
        doc2.Save(doc2Path, SaveFormat.Docx);

        // Load the first document as the destination.
        Document mergedDoc = new Document(doc1Path);

        // Load the second document to be appended.
        Document srcDoc = new Document(doc2Path);

        // Append the second document while preserving its original formatting.
        mergedDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as PDF.
        mergedDoc.Save(mergedPdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(mergedPdfPath))
            throw new InvalidOperationException("The merged PDF file was not created.");

        // Load the PDF back to verify that it contains content from both source documents.
        Document verifyDoc = new Document(mergedPdfPath);
        string mergedText = verifyDoc.GetText();

        if (!mergedText.Contains("This is the first sample document.") ||
            !mergedText.Contains("This is the second sample document."))
        {
            throw new InvalidOperationException("The merged PDF does not contain content from all source documents.");
        }

        // Cleanup: optional removal of intermediate files.
        // File.Delete(doc1Path);
        // File.Delete(doc2Path);
    }
}
