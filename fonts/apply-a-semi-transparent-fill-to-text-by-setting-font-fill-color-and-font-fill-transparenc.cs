using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;

namespace FontFillExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add a paragraph with text.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, semi‑transparent text!");

            // Create an Aspose.Drawing.Color (red) and convert it to System.Drawing.Color.
            Aspose.Drawing.Color asposeColor = Aspose.Drawing.Color.Red;
            System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(asposeColor.ToArgb());

            // Apply the fill color and set transparency (0.0 = opaque, 1.0 = fully transparent).
            builder.Font.Fill.Color = sysColor;
            builder.Font.Fill.Transparency = 0.5; // 50 % transparent

            // Validate that the properties were set correctly.
            double transparency = builder.Font.Fill.Transparency;
            System.Drawing.Color appliedColor = builder.Font.Fill.Color;

            if (Math.Abs(transparency - 0.5) > 0.0001 ||
                appliedColor.ToArgb() != sysColor.ToArgb())
            {
                throw new InvalidOperationException("Font fill properties were not applied correctly.");
            }

            // Save the document to disk.
            string outputPath = "Output.docx";
            doc.Save(outputPath);

            // Ensure the file was created.
            if (!File.Exists(outputPath))
            {
                throw new FileNotFoundException("The output file was not created.", outputPath);
            }
        }
    }
}
