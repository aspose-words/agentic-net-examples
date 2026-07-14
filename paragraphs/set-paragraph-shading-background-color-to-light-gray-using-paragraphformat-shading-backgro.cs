using System;
using System.IO;
using System.Drawing;
using Aspose.Words;

namespace ParagraphShadingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set the paragraph shading background color to light gray.
            builder.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightGray;

            // Write some text to the paragraph.
            builder.Writeln("This paragraph has a light gray background shading.");

            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ParagraphShading.docx");

            // Save the document to the specified path.
            doc.Save(outputPath);
        }
    }
}
