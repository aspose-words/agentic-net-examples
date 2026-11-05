using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the sample documents and the merged PDF.
        string destinationPath = "Destination.docx";
        string sourcePath = "Source.docx";
        string outputPdfPath = "Merged.pdf";

        // -------------------------------------------------------------
        // Create the destination document with a custom paragraph style.
        // -------------------------------------------------------------
        Document destinationDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destinationDoc);

        // Add a custom style named "MyCustomStyle" with specific formatting.
        Style destStyle = destinationDoc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        destStyle.Font.Size = 14;
        destStyle.Font.Bold = true;

        // Apply the custom style and write a paragraph.
        destBuilder.ParagraphFormat.StyleName = "MyCustomStyle";
        destBuilder.Writeln("This paragraph is in the destination document using MyCustomStyle.");

        // Save the destination document as DOCX.
        destinationDoc.Save(destinationPath, SaveFormat.Docx);

        // -------------------------------------------------------------
        // Create the source document with its own version of the same style.
        // -------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // Add a custom style with the same name but different formatting.
        Style srcStyle = sourceDoc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        srcStyle.Font.Size = 12;
        srcStyle.Font.Italic = true;

        // Apply the source custom style and write a paragraph.
        srcBuilder.ParagraphFormat.StyleName = "MyCustomStyle";
        srcBuilder.Writeln("This paragraph is in the source document using its own MyCustomStyle.");

        // Save the source document as DOCX.
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -------------------------------------------------------------
        // Load the saved documents from disk.
        // -------------------------------------------------------------
        Document destDocLoaded = new Document(destinationPath);
        Document srcDocLoaded = new Document(sourcePath);

        // Append the source document to the destination using destination styles.
        destDocLoaded.AppendDocument(srcDocLoaded, ImportFormatMode.UseDestinationStyles);

        // Save the merged result as PDF.
        destDocLoaded.Save(outputPdfPath, SaveFormat.Pdf);

        // -------------------------------------------------------------
        // Validation
        // -------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException($"The PDF file '{outputPdfPath}' was not created.");
        }

        // Ensure that the merged document contains at least two sections (original + appended).
        if (destDocLoaded.Sections.Count < 2)
        {
            throw new InvalidOperationException("The merged document does not contain the expected number of sections.");
        }

        // Optional cleanup of intermediate DOCX files.
        // File.Delete(destinationPath);
        // File.Delete(sourcePath);
    }
}
