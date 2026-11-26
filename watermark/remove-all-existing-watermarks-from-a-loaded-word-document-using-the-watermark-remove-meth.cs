using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace WatermarkRemovalExample
{
    class Program
    {
        static void Main()
        {
            // Create a sample document with a text watermark.
            string sourcePath = "Sample.docx";
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document with a watermark.");
            doc.Watermark.SetText("CONFIDENTIAL");
            doc.Save(sourcePath);

            // Load the document.
            Document loadedDoc = new Document(sourcePath);
            // Remove all watermarks.
            loadedDoc.Watermark.Remove();
            // Save the result.
            string outputPath = "Output.docx";
            loadedDoc.Save(outputPath);

            // Validate that the output file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine("Watermark removed and document saved to " + outputPath);
            }
        }
    }
}
