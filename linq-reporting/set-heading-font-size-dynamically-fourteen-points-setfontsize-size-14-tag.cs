using System;
using Aspose.Words;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Apply the "Heading 1" style to the current paragraph.
            builder.ParagraphFormat.StyleName = "Heading 1";

            // Set the font size of the heading to 14 points.
            builder.Font.Size = 14;

            // Write the heading text.
            builder.Writeln("Dynamic Heading with 14pt Font Size");

            // Save the document to a file.
            doc.Save("HeadingFontSize.docx");
        }
    }
}
