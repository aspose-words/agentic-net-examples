// Path to the folder where the output PDF will be saved.
string outputPath = @"C:\Temp\HyphenatedDocument.pdf";

// Create a new empty document.
Aspose.Words.Document doc = new Aspose.Words.Document();
Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);

// Add some long text that will require hyphenation when wrapped.
builder.Font.Size = 12;
builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

// Enable automatic hyphenation for the whole document.
doc.HyphenationOptions.AutoHyphenation = true;

// Optional: fine‑tune hyphenation behavior.
doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;   // Max two consecutive hyphenated lines.
doc.HyphenationOptions.HyphenationZone = 720;       // 0.5 inch from the right margin.
doc.HyphenationOptions.HyphenateCaps = true;       // Hyphenate words in all caps.

// Save the document as PDF – hyphens will be inserted where needed.
doc.Save(outputPath, Aspose.Words.SaveFormat.Pdf);
