using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add introductory text.
        builder.Writeln("Below is an equation displayed on a separate line:");

        // Insert an EQ field to bootstrap a real OfficeMath node.
        Field field = builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ argument (a simple fraction).
        builder.MoveTo(field.Separator);
        builder.Write(@"\f(1,2)");

        // Update the field so that the equation is parsed.
        field.Update();

        // Convert the field to OfficeMath.
        FieldEQ eqField = field as FieldEQ;
        OfficeMath officeMath = eqField?.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the field.
        Node fieldStart = field.Start;
        CompositeNode parent = fieldStart.ParentNode as CompositeNode;
        if (parent == null)
            throw new InvalidOperationException("Field start does not have a composite parent.");

        parent.InsertBefore(officeMath, fieldStart);
        field.Remove();

        // Configure the OfficeMath to display on a separate line.
        officeMath.DisplayType = OfficeMathDisplayType.Display;

        // Save the document.
        string outputPath = "OfficeMathDisplay.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify the output file exists.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("Output file was not created.", outputPath);

        // Reload and verify the display type.
        Document loaded = new Document(outputPath);
        OfficeMath loadedMath = loaded.GetChildNodes(NodeType.OfficeMath, true)[0] as OfficeMath;
        if (loadedMath == null || loadedMath.DisplayType != OfficeMathDisplayType.Display)
            throw new InvalidOperationException("OfficeMath display type was not set correctly.");
    }
}
