using Aspose.Words;
using Aspose.Words.Saving;

class AddImageWatermark
{
    static void Main()
    {
        // Load the existing document (any format supported by Aspose.Words)
        Document doc = new Document("Input.docx");

        // Configure image watermark appearance
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            // Disable washout to make the watermark fully opaque
            IsWashout = false,
            // Scale factor (optional, 0 = auto). Adjust as needed.
            Scale = 5
        };

        // Add the image watermark from a file.
        // Use the overload that accepts a file path to avoid the System.Drawing dependency.
        doc.Watermark.SetImage("Watermark.png", watermarkOptions);

        // Save the document in WordML (XML) format
        WordML2003SaveOptions saveOptions = new WordML2003SaveOptions();
        doc.Save("Output.xml", saveOptions);
    }
}
