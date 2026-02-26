using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

class InsertOfficeMathExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a heading before the equation.
        builder.Writeln("Example of inserting OfficeMath:");

        // Insert an EQ field with a simple equation (e.g., a fraction a/b).
        // The field code follows the Word EQ field syntax.
        builder.InsertField("EQ \\x a \\y b");

        // Retrieve the last inserted field, which is the EQ field we just added.
        FieldEQ fieldEQ = doc.Range.Fields
            .OfType<FieldEQ>()
            .LastOrDefault();

        if (fieldEQ != null)
        {
            // Convert the EQ field to an OfficeMath object.
            OfficeMath officeMath = fieldEQ.AsOfficeMath();

            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start node.
                fieldEQ.Start.ParentNode.InsertBefore(officeMath, fieldEQ.Start);

                // Remove the original EQ field from the document.
                fieldEQ.Remove();

                // Optionally set the display type (Display = on its own line).
                officeMath.DisplayType = OfficeMathDisplayType.Display;
            }
        }

        // Save the document to a DOCX file.
        doc.Save("OfficeMathInserted.docx");
    }
}
