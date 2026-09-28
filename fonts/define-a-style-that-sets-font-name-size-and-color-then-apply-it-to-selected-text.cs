using System;
using System.IO;
using Aspose.Words;

namespace FontStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some initial text.
            builder.Writeln("This paragraph uses the default style.");

            // Define a new character style.
            Style customStyle = doc.Styles.Add(StyleType.Character, "MyCustomStyle");
            customStyle.Font.Name = "Arial";
            customStyle.Font.Size = 16;

            // Set the font color using Aspose.Drawing.Color and convert to System.Drawing.Color.
            Aspose.Drawing.Color asposeColor = Aspose.Drawing.Color.FromArgb(255, 0, 0); // Red
            System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(asposeColor.ToArgb());
            customStyle.Font.Color = sysColor;

            // Validate that the style properties were set correctly.
            if (customStyle.Font.Name != "Arial" ||
                customStyle.Font.Size != 16 ||
                customStyle.Font.Color.ToArgb() != sysColor.ToArgb())
            {
                throw new InvalidOperationException("Font properties were not set correctly on the style.");
            }

            // Apply the custom style to selected text.
            builder.Font.StyleName = "MyCustomStyle";
            builder.Writeln("This paragraph uses the custom style with Arial, 16pt, red color.");

            // Reset to default style for any further text.
            builder.Font.StyleName = "Default Paragraph Font";

            // Save the document.
            string outputPath = "StyledDocument.docx";
            doc.Save(outputPath);

            // Ensure the file was created.
            if (!File.Exists(outputPath))
            {
                throw new FileNotFoundException("The output document was not created.", outputPath);
            }
        }
    }
}
