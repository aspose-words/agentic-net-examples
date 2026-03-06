using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleIntoHtmlDocument
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some HTML content.
        const string html = "<h2>Embedded Files</h2>" +
                            "<p>This document contains an embedded ZIP archive.</p>";
        builder.InsertHtml(html);

        // Path placeholders – replace with actual paths in your environment.
        string zipFilePath = Path.Combine(MyDir, "sample.zip");          // file to embed
        string iconFilePath = Path.Combine(ImageDir, "icon.ico");       // custom icon (optional)

        // Insert the ZIP file as an OLE object displayed as an icon.
        // The ProgID "Package" tells Aspose.Words to treat the data as a generic OLE package.
        // The icon caption will be "Sample ZIP".
        using (FileStream zipStream = new FileStream(zipFilePath, FileMode.Open, FileAccess.Read))
        {
            builder.InsertOleObjectAsIcon(zipStream, "Package", iconFilePath, "Sample ZIP");
        }

        // Save the resulting document.
        string outputPath = Path.Combine(ArtifactsDir, "HtmlWithOle.docx");
        doc.Save(outputPath);
    }

    // Placeholder directories – set these to appropriate locations before running.
    private static readonly string MyDir = @"C:\Data\";
    private static readonly string ImageDir = @"C:\Data\Images\";
    private static readonly string ArtifactsDir = @"C:\Data\Output\";
}
