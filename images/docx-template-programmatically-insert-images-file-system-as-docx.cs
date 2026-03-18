using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsImageInsertion
{
    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that will be loaded.
            const string templatePath = @"C:\Docs\Template.docx";

            // Load the existing template document.
            Document doc = new Document(templatePath);

            // Create a DocumentBuilder for the loaded document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // List of image file paths to insert into the document.
            List<string> imagePaths = new List<string>
            {
                @"C:\Images\Photo1.jpg",
                @"C:\Images\Diagram.png",
                @"C:\Images\Logo.gif"
            };

            // Insert each image at the current cursor position.
            // Images are inserted as inline shapes with their original dimensions.
            foreach (string imagePath in imagePaths)
            {
                // Insert the image.
                Shape imageShape = builder.InsertImage(imagePath);

                // Optional: add a paragraph break after each image for spacing.
                builder.Writeln();
            }

            // Path where the resulting document will be saved.
            const string outputPath = @"C:\Docs\Result.docx";

            // Save the modified document as DOCX.
            doc.Save(outputPath);
        }
    }
}
