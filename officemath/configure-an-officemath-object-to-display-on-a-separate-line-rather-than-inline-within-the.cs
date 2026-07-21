using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathDisplayExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Writeln("Below is an equation displayed on its own line:");

        // Insert an EQ field that will be converted to a real OfficeMath object.
        // The field code "EQ" is created automatically.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ argument (a simple fraction 1/2).
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(eqField.Start.ParentNode);

        // Update the field to ensure the field code is processed.
        eqField.Update();

        // Convert the EQ field to an OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Verify that we are working with a top‑level equation.
        if (officeMath.MathObjectType != MathObjectType.OMathPara)
            throw new InvalidOperationException("The created OfficeMath is not a top‑level equation.");

        // Set the display type to Display so the equation appears on a separate line.
        officeMath.DisplayType = OfficeMathDisplayType.Display;
        // Set justification (e.g., left aligned). Must be set after DisplayType.
        officeMath.Justification = OfficeMathJustification.Left;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OfficeMathDisplay.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
