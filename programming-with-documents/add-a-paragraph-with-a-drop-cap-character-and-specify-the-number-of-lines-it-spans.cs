using System;
using System.IO;
using Aspose.Words;

namespace DropCapExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Attach a DocumentBuilder to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set the number of lines the drop cap should span.
            // This makes the next paragraph start with a large capital letter.
            builder.ParagraphFormat.LinesToDrop = 4;

            // Write the drop cap character.
            builder.Writeln("H");

            // Reset LinesToDrop to 0 so subsequent text behaves normally.
            builder.ParagraphFormat.LinesToDrop = 0;

            // Write the remaining text that will wrap around the drop cap.
            builder.Writeln("ello world! This paragraph demonstrates a drop cap character spanning four lines.");

            // Define the output path (in the same folder as the executable).
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DropCapExample.docx");

            // Save the document.
            doc.Save(outputPath);
        }
    }
}
