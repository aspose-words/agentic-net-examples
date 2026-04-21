using System;
using System.IO;
using Aspose.Words;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for the source and output documents.
        string sourcePath = "Sample.docx";
        string outputPath = "Watermarked.docx";

        // Create a simple source document if it does not already exist.
        if (!File.Exists(sourcePath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is a sample document for watermark demonstration.");
            sampleDoc.Save(sourcePath);
        }

        // Load the document from the file system.
        Document doc = new Document(sourcePath);

        // Define optional text watermark formatting.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Add a text watermark to the loaded document.
        doc.Watermark.SetText("Confidential", options);

        // Save the watermarked document.
        doc.Save(outputPath);
    }
}
