using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

class ReplaceOfficeMath
{
    static void Main()
    {
        // Load an existing document that contains an OfficeMath object.
        Document doc = new Document("Input.docx");

        // Find the first OfficeMath node that we want to replace.
        OfficeMath oldMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
        if (oldMath == null)
        {
            Console.WriteLine("No OfficeMath object found in the document.");
            return;
        }

        // The new equation expressed in EQ field syntax.
        // Example: a fraction 1 over 4 -> "\f(1,4)"
        string newEquation = @"\f(1,4)";

        // Insert a temporary EQ field that will generate the desired OfficeMath.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the field at the location of the old OfficeMath's parent paragraph.
        // This ensures the new OfficeMath will be created in the same story.
        builder.MoveTo(oldMath.ParentParagraph);
        FieldEQ tempField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the equation arguments into the field separator.
        builder.MoveTo(tempField.Separator);
        builder.Write(newEquation);

        // Move back to the paragraph to keep the document structure tidy.
        builder.MoveTo(tempField.Start.ParentNode);

        // Convert the EQ field to an OfficeMath object.
        OfficeMath newMath = tempField.AsOfficeMath();

        // Insert the new OfficeMath before the old one.
        oldMath.ParentParagraph.InsertBefore(newMath, oldMath);

        // Remove the old OfficeMath and the temporary field.
        oldMath.Remove();
        tempField.Remove();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
