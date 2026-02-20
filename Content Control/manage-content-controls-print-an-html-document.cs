using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("input.docx"); // replace with your file path

        // Configure HTML save options.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Save the document to a memory stream as HTML.
        using (MemoryStream htmlStream = new MemoryStream())
        {
            doc.Save(htmlStream, htmlOptions);

            // Convert the stream contents to a string using the encoding defined in the options (UTF‑8 by default).
            string htmlContent = Encoding.UTF8.GetString(htmlStream.ToArray());

            // Print the HTML to the console.
            Console.WriteLine(htmlContent);
        }
    }
}
