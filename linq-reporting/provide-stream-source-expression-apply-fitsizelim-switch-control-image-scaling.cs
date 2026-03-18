using Aspose.Words;
using Aspose.Words.Drawing;
using System.IO;

// Create a new document and a builder.
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

// Stream source expression for the image.
using (Stream imageStream = File.OpenRead(@"C:\Images\Sample.jpg"))
{
    // Insert the image from the stream as an inline shape.
    Shape shape = builder.InsertImage(imageStream);

    // Apply the fitSizeLim behavior: fit the image data to the shape's frame
    // so that the aspect ratio of the image matches the shape's dimensions.
    shape.ImageData.FitImageToShape();
}

// Save the resulting document.
doc.Save(@"C:\Output\Result.docx");
