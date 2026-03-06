using System;
using System.IO;
using Aspose.Words;

class InsertHtmlFromTemplate
{
    static void Main()
    {
        // Path to the TXT file that contains the HTML fragment.
        string templatePath = @"Template.txt";

        // Read the entire content of the TXT file.
        // The file is expected to hold valid HTML markup.
        string html = File.ReadAllText(templatePath);

        // Create a new blank Word document.
        Document doc = new Document();

        // Obtain a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the HTML string into the document.
        // This parses the HTML and converts it to Word formatting.
        builder.InsertHtml(html);

        // Save the resulting document.
        doc.Save(@"Result.docx");
    }
}
