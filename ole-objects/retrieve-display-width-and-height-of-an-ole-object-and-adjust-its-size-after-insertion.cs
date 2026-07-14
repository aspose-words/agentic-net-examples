using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Ole;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a temporary text file that will be embedded as an OLE object.
        string tempFilePath = Path.Combine(Path.GetTempPath(), "Sample.txt");
        File.WriteAllText(tempFilePath, "Hello Aspose.Words OLE object!");

        // Insert the OLE object into the document.
        // Parameters: file name, isLinked (false = embed), asIcon (false = show content), presentation (null = default).
        Shape oleShape = builder.InsertOleObject(tempFilePath, false, false, null);

        // Retrieve the current display size of the OLE object (in points).
        double originalWidth = oleShape.Width;
        double originalHeight = oleShape.Height;
        Console.WriteLine($"Original OLE object size: {originalWidth} pt x {originalHeight} pt");

        // Adjust the size of the OLE object.
        // Example: increase width by 50% and height by 30%.
        oleShape.Width = originalWidth * 1.5;
        oleShape.Height = originalHeight * 1.3;
        Console.WriteLine($"Adjusted OLE object size: {oleShape.Width} pt x {oleShape.Height} pt");

        // If the OLE object is an ActiveX control, its internal control size can also be set.
        if (oleShape.OleFormat?.OleControl is Forms2OleControl oleControl)
        {
            oleControl.Width = oleShape.Width;
            oleControl.Height = oleShape.Height;
        }

        // Save the resulting document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectDemo.docx");
        doc.Save(outputPath);
    }
}
