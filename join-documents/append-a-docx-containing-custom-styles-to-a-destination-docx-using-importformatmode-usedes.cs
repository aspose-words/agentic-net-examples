using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths.
        string destPath = Path.Combine(outputDir, "Destination.docx");
        string srcPath = Path.Combine(outputDir, "Source.docx");
        string mergedPath = Path.Combine(outputDir, "Merged.docx");
        string pdfPath = Path.Combine(outputDir, "Merged.pdf");

        // ---------- Create destination document with a custom style ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Define a custom style named "CustomStyle".
        Style destStyle = destDoc.Styles.Add(StyleType.Paragraph, "CustomStyle");
        destStyle.Font.Name = "Arial";
        destStyle.Font.Size = 14;
        destStyle.Font.Color = Color.Blue;

        // Apply the style and add some text.
        destBuilder.ParagraphFormat.StyleName = destStyle.Name;
        destBuilder.Writeln("Destination document text with custom style.");

        // Save the destination document.
        destDoc.Save(destPath);

        // ---------- Create source document with a style of the same name ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

        // Define a style with the same name but different formatting.
        Style srcStyle = srcDoc.Styles.Add(StyleType.Paragraph, "CustomStyle");
        srcStyle.Font.Name = "Times New Roman";
        srcStyle.Font.Size = 16;
        srcStyle.Font.Color = Color.Red;

        // Apply the style and add some text.
        srcBuilder.ParagraphFormat.StyleName = srcStyle.Name;
        srcBuilder.Writeln("Source document text with custom style.");

        // Save the source document.
        srcDoc.Save(srcPath);

        // ---------- Append source to destination using UseDestinationStyles ----------
        // Load the previously saved documents (optional, can reuse objects).
        Document destination = new Document(destPath);
        Document source = new Document(srcPath);

        // Append while preserving destination styles.
        destination.AppendDocument(source, ImportFormatMode.UseDestinationStyles);

        // Save the merged document.
        destination.Save(mergedPath);

        // ---------- Validate that the merged document contains both texts ----------
        Document mergedDoc = new Document(mergedPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("Destination document text") ||
            !mergedText.Contains("Source document text"))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }

        // ---------- Export merged document to PDF ----------
        mergedDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validate output files exist.
        if (!File.Exists(mergedPath))
            throw new FileNotFoundException("Merged DOCX was not created.", mergedPath);
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF conversion failed.", pdfPath);

        // Inform the user (no interactive input required).
        Console.WriteLine("Documents merged and saved successfully:");
        Console.WriteLine($"- Merged DOCX: {mergedPath}");
        Console.WriteLine($"- PDF: {pdfPath}");
    }
}
