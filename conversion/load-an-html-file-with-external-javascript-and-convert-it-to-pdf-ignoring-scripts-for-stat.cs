using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary files.
        const string htmlFile = "input.html";
        const string jsFile = "script.js";
        const string pdfFile = "output.pdf";

        // Create a simple JavaScript file.
        File.WriteAllText(jsFile, @"
function greet() {
    console.log('Hello from JavaScript!');
}
greet();
");

        // Create an HTML file that references the external JavaScript.
        string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset=""utf-8"">
    <title>Sample HTML with JavaScript</title>
    <script src=""{jsFile}""></script>
</head>
<body>
    <h1>Static Rendering Test</h1>
    <p>This paragraph should appear in the PDF. The script is ignored by Aspose.Words.</p>
</body>
</html>";

        File.WriteAllText(htmlFile, htmlContent);

        // Load the HTML document. Aspose.Words ignores scripts and renders static content only.
        Document doc = new Document(htmlFile);

        // Convert the document to PDF.
        doc.Save(pdfFile, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfFile))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optional cleanup of temporary files (comment out if you need to inspect them).
        // File.Delete(htmlFile);
        // File.Delete(jsFile);
    }
}
