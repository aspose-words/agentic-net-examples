using System;
using System.IO;
using System.Runtime.InteropServices;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define the ProgId to validate.
        string progId = "Word.Document";

        // Validate the ProgId by attempting to get its COM type.
        bool isProgIdValid = false;
        try
        {
            Type comType = Type.GetTypeFromProgID(progId);
            isProgIdValid = comType != null;
        }
        catch (COMException)
        {
            isProgIdValid = false;
        }

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        if (isProgIdValid)
        {
            // Create a temporary text file to embed as an OLE object.
            string tempFilePath = Path.Combine(Path.GetTempPath(), "SampleText.txt");
            File.WriteAllText(tempFilePath, "This is sample text for OLE embedding.");

            // Insert the OLE object using the validated ProgId.
            // Parameters: filePath, progId, isLink (false), isIcon (false), iconCaption (null)
            builder.InsertOleObject(tempFilePath, progId, false, false, null);

            // Clean up the temporary file.
            File.Delete(tempFilePath);
        }
        else
        {
            // If the ProgId is invalid, write a note into the document.
            builder.Writeln($"The ProgId \"{progId}\" is not valid on this system. OLE object was not inserted.");
        }

        // Save the document to the output file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);
    }
}
