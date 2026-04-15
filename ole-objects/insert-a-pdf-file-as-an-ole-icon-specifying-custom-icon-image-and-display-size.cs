using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Prepare a dummy PDF file to embed.
        // This ensures the file exists regardless of the environment.
        // -----------------------------------------------------------------
        string pdfPath = Path.Combine(Environment.CurrentDirectory, "Dummy.pdf");
        if (!File.Exists(pdfPath))
        {
            Document tempPdf = new Document();
            DocumentBuilder tempBuilder = new DocumentBuilder(tempPdf);
            tempBuilder.Writeln("This is a placeholder PDF file.");
            tempPdf.Save(pdfPath, SaveFormat.Pdf);
        }

        // Insert the PDF as an OLE object displayed as an icon.
        // Passing null for the icon file makes Aspose.Words use a predefined default icon.
        Shape oleShape = builder.InsertOleObjectAsIcon(pdfPath, false, null, "My PDF Document");

        // Optionally set the display size of the icon (width and height are in points).
        oleShape.Width = 100;   // 100 points ≈ 1.39 inches
        oleShape.Height = 100;  // 100 points ≈ 1.39 inches

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OlePdfIcon.docx");
        doc.Save(outputPath);
    }
}
