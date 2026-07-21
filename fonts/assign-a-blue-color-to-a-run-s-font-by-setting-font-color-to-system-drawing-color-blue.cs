using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a System.Drawing.Color instance from Aspose.Drawing.Color.Blue.
        System.Drawing.Color blueColor = System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Blue.ToArgb());

        // Assign the blue color to the current font (affects the Run that will be inserted).
        builder.Font.Color = blueColor;

        // Insert a run of text.
        builder.Writeln("Hello World!");

        // Define output path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "BlueRun.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully to: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
