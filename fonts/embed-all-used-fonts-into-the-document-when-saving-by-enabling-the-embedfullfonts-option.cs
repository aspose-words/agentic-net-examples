using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory and file path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "EmbeddedFonts.pdf");

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with Arial font.
        builder.Font.Name = "Arial";
        builder.Font.Size = 24;
        // Use Aspose.Drawing.Color and convert to System.Drawing.Color.
        Aspose.Drawing.Color aspColor1 = Aspose.Drawing.Color.Blue;
        System.Drawing.Color sysColor1 = System.Drawing.Color.FromArgb(aspColor1.ToArgb());
        builder.Font.Color = sysColor1;
        builder.Writeln("This text uses Arial.");

        // Validate the first font settings.
        if (builder.Font.Name != "Arial")
            throw new InvalidOperationException("Font name was not set to Arial.");

        // Second paragraph with Times New Roman font.
        builder.Font.Name = "Times New Roman";
        builder.Font.Size = 18;
        Aspose.Drawing.Color aspColor2 = Aspose.Drawing.Color.Green;
        System.Drawing.Color sysColor2 = System.Drawing.Color.FromArgb(aspColor2.ToArgb());
        builder.Font.Color = sysColor2;
        builder.Writeln("This text uses Times New Roman.");

        // Validate the second font settings.
        if (builder.Font.Name != "Times New Roman")
            throw new InvalidOperationException("Font name was not set to Times New Roman.");

        // Configure PDF save options to embed all fonts fully.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };

        // Save the document as PDF.
        doc.Save(pdfPath, saveOptions);

        // Ensure the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("The PDF file was not created.", pdfPath);

        // Indicate successful completion.
        Console.WriteLine("PDF saved with full font embedding at: " + pdfPath);
    }
}
