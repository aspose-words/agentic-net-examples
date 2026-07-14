using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add introductory text.
        builder.Writeln("Below is an equation displayed on its own line:");

        // Insert an EQ field (the field code initially contains "EQ").
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write a simple fraction equation.
        // This appends the EQ switch arguments after the "EQ" code.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)"); // 1 over 2

        // Update the field so that the EQ code is recognized.
        eqField.Update();

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(eqField.Start.ParentNode);

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Configure the OfficeMath to display on a separate line.
        // This should be done only on top‑level OfficeMath (MathObjectType.OMathPara).
        officeMath.DisplayType = OfficeMathDisplayType.Display;
        officeMath.Justification = OfficeMathJustification.Left;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
