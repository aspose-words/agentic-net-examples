using System;
using System.IO;
using Aspose.Words;

class InsertOleIntoMhtml
{
    static void Main()
    {
        // Path to the source MHTML document.
        string mhtmlPath = @"C:\Docs\SourceDocument.mht";

        // Path to the file that will be embedded as an OLE object (e.g., a PDF).
        string oleFilePath = @"C:\Docs\Sample.pdf";

        // Optional: path to an icon image to represent the OLE object.
        // If null, Aspose.Words will use a predefined icon.
        string iconPath = @"C:\Images\PdfIcon.png";

        // Load the MHTML document.
        Document doc = new Document(mhtmlPath);

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph before the OLE object for context.
        builder.Writeln("Embedded PDF as OLE object:");

        // Open the icon image as a stream (if an icon is desired).
        Stream iconStream = null;
        if (File.Exists(iconPath))
        {
            iconStream = new FileStream(iconPath, FileMode.Open, FileAccess.Read);
        }

        // Insert the OLE object.
        // Parameters: file name, isLinked (false = embed), asIcon (true to display as icon), presentation stream.
        builder.InsertOleObject(oleFilePath, false, true, iconStream);

        // Clean up the icon stream if it was opened.
        if (iconStream != null)
        {
            iconStream.Dispose();
        }

        // Save the modified document. Here we save as DOCX, but you can also save back to MHTML.
        string outputPath = @"C:\Docs\ResultDocument.docx";
        doc.Save(outputPath);
    }
}
