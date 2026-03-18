using System;
using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph shading background color to light gray.
        builder.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightGray;

        // Write a paragraph to demonstrate the shading.
        builder.Writeln("This paragraph has a light gray background shading.");

        // Save the document to a file.
        string outputFile = "ParagraphShading.docx";
        doc.Save(outputFile);
    }
}
