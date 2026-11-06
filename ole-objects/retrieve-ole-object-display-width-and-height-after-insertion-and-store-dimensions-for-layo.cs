using System;
using System.Runtime.InteropServices;

public class Program
{
    public static void Main()
    {
        // Create Word application via COM late binding
        Type wordAppType = Type.GetTypeFromProgID("Word.Application");
        if (wordAppType == null)
        {
            Console.WriteLine("Word is not installed on this machine.");
            return;
        }

        object wordAppObj = Activator.CreateInstance(wordAppType);
        dynamic wordApp = wordAppObj;
        wordApp.Visible = false;

        // Add a new document
        object docObj = wordApp.Documents.Add();
        dynamic doc = docObj;

        // Insert an OLE object (Excel worksheet) at the start of the document
        object classType = "Excel.Sheet";
        object fileName = "";
        object linkToFile = false;
        object displayAsIcon = false;
        object iconFileName = Type.Missing;
        object iconIndex = Type.Missing;
        object iconLabel = Type.Missing;
        object range = doc.Range(0, 0);

        object oleShapeObj = doc.InlineShapes.AddOLEObject(
            classType,
            fileName,
            linkToFile,
            displayAsIcon,
            iconFileName,
            iconIndex,
            iconLabel,
            range);
        dynamic oleShape = oleShapeObj;

        // Retrieve display dimensions (points)
        float width = (float)oleShape.Width;
        float height = (float)oleShape.Height;

        // Example layout calculation: area
        float area = width * height;

        // Output dimensions
        Console.WriteLine($"OLE object width: {width} points");
        Console.WriteLine($"OLE object height: {height} points");
        Console.WriteLine($"Calculated area: {area} square points");

        // Clean up
        doc.Close(false);
        wordApp.Quit();

        // Release COM objects
        Marshal.FinalReleaseComObject(oleShapeObj);
        Marshal.FinalReleaseComObject(docObj);
        Marshal.FinalReleaseComObject(wordAppObj);
    }
}
