using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Define deterministic file names.
        string workDir = Directory.GetCurrentDirectory();
        string imagePath = Path.Combine(workDir, "sample.png");
        string docPath = Path.Combine(workDir, "sample.docx");
        string htmlPath = Path.Combine(workDir, "output.html");

        // -------------------------------------------------
        // 1. Create a sample PNG image using Aspose.Drawing.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                g.Clear(Color.White);
                // Optionally draw a simple rectangle.
                g.DrawRectangle(Pens.Black, 10, 10, imgWidth - 20, imgHeight - 20);
            }
            // Save the image to a deterministic file.
            bitmap.Save(imagePath);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        // Save the document to a deterministic file.
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the saved document.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // -------------------------------------------------
        // 4. Save the document as HTML with images embedded as Base64.
        // -------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportImagesAsBase64 = true,
            PrettyFormat = true
        };
        loadedDoc.Save(htmlPath, htmlOptions);

        // -------------------------------------------------
        // 5. Validate that the HTML file was created.
        // -------------------------------------------------
        if (!File.Exists(htmlPath))
        {
            throw new FileNotFoundException("HTML output file was not created.", htmlPath);
        }

        // The program finishes here without waiting for user input.
    }
}
