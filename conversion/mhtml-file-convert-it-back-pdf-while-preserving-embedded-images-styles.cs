using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class MhtmlToPdfConverter
{
    static void Main()
    {
        // Create a temporary MHTML file with minimal content.
        string tempMhtmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".mht");
        File.WriteAllText(tempMhtmlPath,
            @"MIME-Version: 1.0
Content-Type: multipart/related; boundary=""----=_NextPart_000_0000""; type=""text/html""

------=_NextPart_000_0000
Content-Type: text/html; charset=""utf-8""

<html><body><h1>Hello, World!</h1><p>This is a test MHTML document.</p></body></html>

------=_NextPart_000_0000--");

        // Path where the resulting PDF will be saved (in the current directory).
        string outputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "result.pdf");

        // Load options for HTML/MHTML files.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            LoadFormat = LoadFormat.Mhtml
        };

        // Load the MHTML document into an Aspose.Words Document object.
        Document document = new Document(tempMhtmlPath, loadOptions);

        // Save the document as PDF.
        document.Save(outputPdfPath, SaveFormat.Pdf);

        // Clean up the temporary MHTML file.
        File.Delete(tempMhtmlPath);

        Console.WriteLine($"PDF saved to: {outputPdfPath}");
    }
}
