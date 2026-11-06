using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

namespace FontFillExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a semi‑transparent fill color using Aspose.Drawing.Color.
            Aspose.Drawing.Color asposeColor = Aspose.Drawing.Color.FromArgb(255, 0, 0); // Red.

            // Convert Aspose.Drawing.Color to System.Drawing.Color because Font.Fill.Color expects System.Drawing.Color.
            builder.Font.Fill.Color = System.Drawing.Color.FromArgb(asposeColor.ToArgb());

            // Set transparency (0.0 = opaque, 1.0 = fully transparent).
            builder.Font.Fill.Transparency = 0.5; // 50% transparent.

            // Write a line of text with the specified fill.
            builder.Writeln("Semi‑transparent filled text");

            // Validate that the properties were applied.
            var fill = builder.Font.Fill;
            Console.WriteLine("Fill Color ARGB: " + fill.Color.ToArgb());
            Console.WriteLine("Transparency: " + fill.Transparency);

            // Save the document to disk.
            string outputPath = "Output.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine("Document saved successfully: " + Path.GetFullPath(outputPath));
            }
            else
            {
                Console.WriteLine("Failed to save the document.");
            }
        }
    }
}
