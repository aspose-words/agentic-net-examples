using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class SetOfficeMathDisplayInline
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Introductory paragraph.
        builder.Writeln("Below is an equation that will be set to inline display mode:");

        // Insert an EQ field (the placeholder for a real OfficeMath object).
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ eqField = field as FieldEQ;
        if (eqField == null)
            throw new InvalidOperationException("Failed to create an EQ field.");

        // Write the EQ switch/arguments after the field separator.
        // The field code will become: EQ \f(1,2)
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");

        // Return the builder to the paragraph that contains the field and start a new line.
        builder.MoveTo(eqField.Start.ParentNode);
        builder.Writeln();

        // Ensure the field is up‑to‑date (optional but safe).
        eqField.Update();

        // Convert the EQ field to a real OfficeMath node.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Set the display type of all top‑level OfficeMath paragraphs to Inline.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                om.DisplayType = OfficeMathDisplayType.Inline;
        }

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OfficeMathInline.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
