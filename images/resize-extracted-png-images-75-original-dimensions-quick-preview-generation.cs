using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;

class ImagePreviewGenerator
{
    static void Main()
    {
        // Paths for the source document and the folder where previews will be saved.
        string sourceDocPath = @"C:\Docs\input.docx";
        string previewFolder = @"C:\Previews\";

        // Load the Word document.
        Document doc = new Document(sourceDocPath);

        // Counter to generate unique file names for each extracted image.
        int imageIndex = 0;

        // Iterate over all Shape nodes in the document (including those inside headers/footers).
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Process only shapes that actually contain an image.
            if (shape.HasImage)
            {
                // Configure image save options to scale the output to 75% of the original size.
                ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
                {
                    Scale = 0.75f // 75% zoom factor.
                };

                // Render the shape (image) to a PNG file using the specified scaling.
                shape.GetShapeRenderer().Save($"{previewFolder}image_{imageIndex}.png", saveOptions);
                imageIndex++;
            }
        }
    }
}
