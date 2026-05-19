using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for the example files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeOleExample");
        Directory.CreateDirectory(tempFolder);

        // Paths for the PDF file and the output document.
        string pdfPath = Path.Combine(tempFolder, "Sample.pdf");
        string outputDocPath = Path.Combine(tempFolder, "Result.docx");

        // -----------------------------------------------------------------
        // 1) Create a simple PDF file that will be embedded as an OLE object.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document();
        DocumentBuilder pdfBuilder = new DocumentBuilder(pdfDoc);
        pdfBuilder.Writeln("This is a sample PDF file embedded as an OLE icon.");
        pdfDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 2) Create the main document and insert the PDF as an OLE icon.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the PDF file as an embedded OLE object displayed as an icon.
        // Use the overload without progId – Word will detect the file type automatically.
        // Pass null for the custom icon file to let Aspose.Words use the default icon.
        Shape oleShape = builder.InsertOleObjectAsIcon(
            pdfPath,          // fileName
            false,            // isLinked – embed the file
            null,             // iconFile – use default icon
            "Open Sample PDF" // iconCaption
        );

        // Optionally set the display size of the icon (e.g., 50 mm x 50 mm).
        oleShape.Width = ConvertUtil.MillimeterToPoint(50);
        oleShape.Height = ConvertUtil.MillimeterToPoint(50);

        // -----------------------------------------------------------------
        // 3) Save the resulting document.
        // -----------------------------------------------------------------
        doc.Save(outputDocPath, SaveFormat.Docx);
    }
}
