using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfHyphenationToImage
{
    static void Main()
    {
        // Create a new document with some long text to demonstrate hyphenation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("This is a demonstration of automatic hyphenation in a long word: extraordinaryextraordinaryextraordinary.");

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Prepare image save options.
        ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            PageSet = new PageSet(0), // first page
            Resolution = 150
        };

        // Save the rendered page to an image file in the current directory.
        string imageFile = "Rendered.png";
        doc.Save(imageFile, imgOptions);

        Console.WriteLine($"Image saved to {System.IO.Path.GetFullPath(imageFile)}");
    }
}
