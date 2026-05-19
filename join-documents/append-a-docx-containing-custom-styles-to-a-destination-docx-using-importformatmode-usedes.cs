using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Define output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the sample documents
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        string destinationPath = Path.Combine(outputDir, "Destination.docx");
        string mergedPath = Path.Combine(outputDir, "Merged.docx");
        string pdfPath = Path.Combine(outputDir, "Merged.pdf");

        // -------------------------------------------------
        // Create source document with a custom paragraph style
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // Define a custom style named "MyCustomStyle"
        Style customStyle = sourceDoc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Name = "Courier New";
        customStyle.Font.Size = 14;
        customStyle.Font.Color = Color.Blue;

        // Apply the custom style to a paragraph
        srcBuilder.ParagraphFormat.StyleName = "MyCustomStyle";
        srcBuilder.Writeln("This paragraph uses a custom style defined in the source document.");

        // Save the source document (DOCX)
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -------------------------------------------------
        // Create destination document (plain content)
        // -------------------------------------------------
        Document destinationDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(destinationDoc);
        dstBuilder.Writeln("This is the beginning of the destination document.");

        // Save the destination document (optional, just for completeness)
        destinationDoc.Save(destinationPath, SaveFormat.Docx);

        // -------------------------------------------------
        // Append the source document to the destination document
        // using ImportFormatMode.UseDestinationStyles
        // -------------------------------------------------
        destinationDoc.AppendDocument(sourceDoc, ImportFormatMode.UseDestinationStyles);

        // Save the merged document
        destinationDoc.Save(mergedPath, SaveFormat.Docx);

        // -------------------------------------------------
        // Export the merged document to PDF
        // -------------------------------------------------
        destinationDoc.Save(pdfPath, SaveFormat.Pdf);

        // -------------------------------------------------
        // Simple validation: ensure files were created
        // -------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new FileNotFoundException("Merged DOCX was not created.", mergedPath);
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF conversion failed.", pdfPath);

        // Optional content check: merged document should contain text from both parts
        string mergedText = destinationDoc.GetText();
        if (!mergedText.Contains("This is the beginning of the destination document.") ||
            !mergedText.Contains("This paragraph uses a custom style defined in the source document."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }

        // Program completes without interactive prompts
    }
}
