using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertPdfAsOleIcon
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paths to the PDF file to embed and the custom ICO file for the icon.
        string pdfFilePath = @"C:\Files\Sample.pdf";
        string iconFilePath = @"C:\Files\CustomIcon.ico";

        // Insert the PDF as an OLE object displayed as an icon.
        // Parameters:
        //   fileName   – full path to the PDF file.
        //   isLinked   – false = embed the PDF, true = link to the PDF.
        //   iconFile   – full path to the .ico file that will be used as the icon.
        //   iconCaption– caption displayed under the icon (null uses the file name).
        Shape oleShape = builder.InsertOleObjectAsIcon(pdfFilePath, false, iconFilePath, "My PDF Document");

        // Adjust the size of the icon (width and height are in points).
        oleShape.Width = 100;   // 100 points ≈ 1.39 inches
        oleShape.Height = 100;  // 100 points ≈ 1.39 inches

        // Save the document.
        doc.Save(@"C:\Files\Result.docx");
    }
}
