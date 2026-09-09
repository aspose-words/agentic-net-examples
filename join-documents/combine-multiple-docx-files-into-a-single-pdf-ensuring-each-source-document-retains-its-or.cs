using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define a folder for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Paths for the sample source documents.
        string doc1Path = Path.Combine(workDir, "Source1.docx");
        string doc2Path = Path.Combine(workDir, "Source2.docx");

        // Create first sample DOCX with a heading and some text.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder1.Writeln("First Document Heading");
        builder1.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder1.Font.Color = System.Drawing.Color.Blue;
        builder1.Writeln("This is the first sample document.");
        doc1.Save(doc1Path, SaveFormat.Docx);

        // Create second sample DOCX with different formatting.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder2.Writeln("Second Document Subheading");
        builder2.Font.Color = System.Drawing.Color.Green;
        builder2.Writeln("Content of the second sample document goes here.");
        doc2.Save(doc2Path, SaveFormat.Docx);

        // Load the source documents.
        Document srcDoc1 = new Document(doc1Path);
        Document srcDoc2 = new Document(doc2Path);

        // Destination document that will hold the merged content.
        Document mergedDoc = new Document();

        // Append the first source document, preserving its original formatting.
        mergedDoc.AppendDocument(srcDoc1, ImportFormatMode.KeepSourceFormatting);

        // Insert a page break between documents for clarity.
        DocumentBuilder mergedBuilder = new DocumentBuilder(mergedDoc);
        mergedBuilder.InsertBreak(BreakType.PageBreak);

        // Append the second source document, also preserving formatting.
        mergedDoc.AppendDocument(srcDoc2, ImportFormatMode.KeepSourceFormatting);

        // Path for the final PDF output.
        string outputPdfPath = Path.Combine(workDir, "MergedOutput.pdf");

        // Save the merged document as PDF.
        mergedDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("The merged PDF file was not created.");
        }

        // Optional: clean up temporary DOCX files (comment out if you need to inspect them).
        // File.Delete(doc1Path);
        // File.Delete(doc2Path);
    }
}
