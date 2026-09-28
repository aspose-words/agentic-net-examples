using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Hello world! This is a sample DOCX document.");
        string docxPath = "sample.docx";
        sampleDoc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document(docxPath);

        // Configure HTML save options. By default, external scripts are not exported,
        // which satisfies the security requirement.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Save as HTML.
        string htmlPath = "output.html";
        doc.Save(htmlPath, htmlOptions);

        // Validate that the HTML file was created and contains data.
        if (!File.Exists(htmlPath))
        {
            throw new InvalidOperationException("The HTML output file was not created.");
        }

        FileInfo info = new FileInfo(htmlPath);
        if (info.Length == 0)
        {
            throw new InvalidOperationException("The HTML output file is empty.");
        }

        // Optionally, clean up the sample DOCX file.
        // File.Delete(docxPath);
    }
}
