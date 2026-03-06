using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the file that will be embedded as an OLE object (e.g., an Excel workbook).
        string oleFilePath = @"C:\Temp\Sample.xlsx";

        // Insert the OLE object as an embedded object (displayed as its content, not as an icon).
        builder.Writeln("Embedded Excel OLE object:");
        Shape oleShape = builder.InsertOleObject(oleFilePath, false, false, null);

        // Insert another OLE object, this time displayed as an icon with a custom icon file and caption.
        string iconPath = @"C:\Temp\CustomIcon.ico";
        builder.InsertParagraph();
        builder.Writeln("Excel OLE object displayed as an icon:");
        Shape oleIconShape = builder.InsertOleObjectAsIcon(oleFilePath, false, iconPath, "My Excel Data");

        // Save the document containing the OLE objects.
        string docPath = @"C:\Temp\OleDemo.docx";
        doc.Save(docPath);

        // Extract the embedded OLE data from the first shape to a separate file.
        OleFormat oleFormat = oleShape.OleFormat;
        string extractedPath = @"C:\Temp\Extracted" + oleFormat.SuggestedExtension;
        oleFormat.Save(extractedPath);
    }
}
