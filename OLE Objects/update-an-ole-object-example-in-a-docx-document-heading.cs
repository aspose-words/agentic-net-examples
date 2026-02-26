using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleObjectExample
{
    static void Main()
    {
        // Path to the folder that contains the source files and where the output will be saved.
        // Adjust these paths to match your environment.
        string dataDir = @"C:\Data\";

        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a heading for the example.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("OLE Object Example");

        // -----------------------------------------------------------------
        // Insert an embedded Excel spreadsheet as a regular OLE object.
        // -----------------------------------------------------------------
        builder.InsertParagraph(); // Add some space.
        builder.Writeln("Embedded Excel:");
        // Parameters: file name, isLinked (false = embed), asIcon (false = show content), presentation (null = default icon).
        builder.InsertOleObject(dataDir + "Spreadsheet.xlsx", false, false, null);

        // -----------------------------------------------------------------
        // Insert a PowerPoint presentation as an OLE object displayed as an icon.
        // -----------------------------------------------------------------
        builder.InsertParagraph(); // Add some space.
        builder.Writeln("PowerPoint as Icon:");
        // Parameters: file name, isLinked (false = embed), icon file, icon caption.
        builder.InsertOleObjectAsIcon(dataDir + "Presentation.pptx", false, dataDir + "Icon.ico", "Presentation");

        // If the document contains fields (e.g., TOC), update them now.
        doc.UpdateFields();

        // Save the document to the specified location.
        doc.Save(dataDir + "OleObjectExample.docx");
    }
}
