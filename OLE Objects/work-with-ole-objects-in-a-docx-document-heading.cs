using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleObjectDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1. Insert an embedded Excel spreadsheet as a regular OLE object.
        // -----------------------------------------------------------------
        // The file path to the Excel file to embed.
        string excelPath = @"MyDir\SampleSpreadsheet.xlsx";

        // Insert the OLE object. Parameters:
        //   fileName   – path to the source file.
        //   isLinked   – false for embedded object.
        //   asIcon     – false to display the content, not an icon.
        //   presentation – null to use the default preview image.
        builder.InsertOleObject(excelPath, false, false, null);

        // Add a paragraph break after the inserted object.
        builder.Writeln();

        // -----------------------------------------------------------------
        // 2. Insert a PowerPoint presentation as an OLE object displayed as an icon.
        // -----------------------------------------------------------------
        string pptPath = @"MyDir\SamplePresentation.pptx";
        string iconPath = @"ImageDir\CustomIcon.ico";
        string iconCaption = "Click to open presentation";

        // Insert the OLE object as an icon with a custom icon file and caption.
        // Parameters:
        //   fileName   – path to the source file.
        //   isLinked   – false for embedded object.
        //   iconFile   – path to the .ico file (custom icon).
        //   iconCaption– caption displayed under the icon.
        builder.InsertOleObjectAsIcon(pptPath, false, iconPath, iconCaption);

        // Add a line break after the icon.
        builder.InsertBreak(BreakType.LineBreak);

        // -----------------------------------------------------------------
        // 3. Extract the first embedded OLE object (the Excel file) to the file system.
        // -----------------------------------------------------------------
        // Retrieve the first shape that contains an OLE object.
        Shape oleShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

        // Access the OleFormat of the shape.
        OleFormat oleFormat = oleShape.OleFormat;

        // Determine a suitable file name using the suggested extension.
        string extractedFileName = Path.Combine(
            @"ArtifactsDir",
            "ExtractedExcel" + oleFormat.SuggestedExtension);

        // Save the embedded OLE data to a file.
        oleFormat.Save(extractedFileName);

        // -----------------------------------------------------------------
        // 4. Save the document containing the OLE objects.
        // -----------------------------------------------------------------
        doc.Save(@"ArtifactsDir\OleObjectsDemo.docx");
    }
}
