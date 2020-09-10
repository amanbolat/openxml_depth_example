using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class Program {
  private static void Main(string[] args) {
    using var doc = WordprocessingDocument.Open("./testdoc.docx", false);
    var reader = OpenXmlReader.Create(doc.MainDocumentPart.Document.Body);

    while (reader.Read()) {
      if (reader.IsEndElement) {
        continue;
      }
      switch (reader.ElementType) {
        case { } typ when typ == typeof(Paragraph) || typ == typeof(Body):
          Console.WriteLine("Name: {0}. Depth: {1}", reader.LocalName, reader.Depth);
          break;
      }
    }
  }
}