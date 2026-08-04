using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current working directory.
        string workDir = Directory.GetCurrentDirectory();
        string coverPath = Path.Combine(workDir, "cover.png");
        string docPath = Path.Combine(workDir, "sample.docx");
        string pdfPath = Path.Combine(workDir, "sample.pdf");

        // Create a simple cover image using Aspose.Drawing.
        CreateCoverImage(coverPath);

        // Create a DOCX document and insert the cover image.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.InsertImage(coverPath);
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the main document content after the cover page.");

        // Save the DOCX file (lifecycle: create → save).
        source.Save(docPath, SaveFormat.Docx);

        // Load the DOCX and convert it to PDF (lifecycle: load → save).
        Document doc = new Document(docPath);
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");
    }

    private static void CreateCoverImage(string filePath)
    {
        // Create a bitmap of size 600x800.
        using (Bitmap bitmap = new Bitmap(600, 800))
        {
            // Obtain a graphics object to draw on the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill the background with a light blue color.
                graphics.Clear(Color.LightBlue);

                // Prepare a drawing font (explicit type to avoid ambiguity).
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48);
                try
                {
                    // Use a solid brush for the text color.
                    using (SolidBrush brush = new SolidBrush(Color.DarkBlue))
                    {
                        // Define the rectangle where the text will be drawn.
                        RectangleF layout = new RectangleF(100, 350, 400, 100);
                        // Draw the text "Cover Page" within the rectangle.
                        graphics.DrawString("Cover Page", font, brush, layout);
                    }
                }
                finally
                {
                    // Ensure the font is disposed.
                    font.Dispose();
                }
            }

            // Save the bitmap as a PNG file using Aspose.Drawing.Imaging.ImageFormat.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
