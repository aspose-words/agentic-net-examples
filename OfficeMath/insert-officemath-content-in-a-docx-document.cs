using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field that represents a simple equation (x = y).
        // The field code "EQ \\x \\y" creates an equation with the characters x and y.
        builder.InsertField("EQ \\x \\y");

        // Locate the inserted EQ field in the document.
        FieldEQ eqField = doc.Range.Fields.OfType<FieldEQ>().FirstOrDefault();

        if (eqField != null)
        {
            // Convert the EQ field to an OfficeMath object.
            OfficeMath officeMath = eqField.AsOfficeMath();

            // Insert the OfficeMath node before the field's start node.
            eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);

            // Remove the original EQ field from the document.
            eqField.Remove();

            // Set the OfficeMath display type to appear on its own line.
            officeMath.DisplayType = OfficeMathDisplayType.Display;
        }

        // Save the resulting document with the inserted OfficeMath.
        doc.Save("OfficeMathInserted.docx");
    }
}
