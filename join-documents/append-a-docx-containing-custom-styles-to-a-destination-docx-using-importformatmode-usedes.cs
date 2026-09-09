using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths for the sample documents and the results.
        string destPath = Path.Combine(outputDir, "Destination.docx");
        string srcPath = Path.Combine(outputDir, "Source.docx");
        string mergedPath = Path.Combine(outputDir, "Merged.docx");
        string pdfPath = Path.Combine(outputDir, "Merged.pdf");

        // ---------- Create destination document ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Add a custom style to the destination document.
        Style destStyle = destDoc.Styles.Add(StyleType.Paragraph, "DestStyle");
        destStyle.Font.Name = "Arial";
        destStyle.Font.Size = 14;
        destStyle.Font.Color = System.Drawing.Color.Blue;

        // Write some text using the custom style.
        destBuilder.ParagraphFormat.StyleName = destStyle.Name;
        destBuilder.Writeln("This is text from the destination document.");

        // Save the destination document.
        destDoc.Save(destPath);

        // ---------- Create source document ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

        // Add a different custom style with the same name to the source document.
        Style srcStyle = srcDoc.Styles.Add(StyleType.Paragraph, "DestStyle");
        srcStyle.Font.Name = "Times New Roman";
        srcStyle.Font.Size = 16;
        srcStyle.Font.Color = System.Drawing.Color.Red;

        // Write some text using the source's custom style.
        srcBuilder.ParagraphFormat.StyleName = srcStyle.Name;
        srcBuilder.Writeln("This is text from the source document.");

        // Save the source document.
        srcDoc.Save(srcPath);

        // ---------- Append source to destination using UseDestinationStyles ----------
        // Load the previously saved documents (optional, we already have them in memory).
        Document destination = new Document(destPath);
        Document source = new Document(srcPath);

        // Append the source document; styles with the same name will adopt the destination's definition.
        destination.AppendDocument(source, ImportFormatMode.UseDestinationStyles);

        // Save the merged document.
        destination.Save(mergedPath);

        // Export the merged document to PDF.
        destination.Save(pdfPath, SaveFormat.Pdf);

        // ---------- Validation ----------
        if (!File.Exists(mergedPath))
            throw new FileNotFoundException("Merged DOCX was not created.", mergedPath);
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF output was not created.", pdfPath);

        // Verify that both pieces of text are present in the merged document.
        string mergedText = destination.GetText();
        if (!mergedText.Contains("This is text from the destination document.") ||
            !mergedText.Contains("This is text from the source document."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }

        // Program completed successfully.
    }
}
