using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ListingAutoFilling
{
    public class Program
    {
        private const string DocFileName = @"C:\university\tst.docx";
        public static void Main(string[] args)
        {
            using WordprocessingDocument wordDoc = WordprocessingDocument.Open(DocFileName, true);
            MainDocumentPart mainPart = wordDoc.MainDocumentPart ?? wordDoc.AddMainDocumentPart();
            
            Body body = mainPart.Document.AppendChild(new Body());

            Run run = new Run();
            RunProperties runProps = new RunProperties();
            Bold bold = new Bold();
            runProps.Append(bold);
            run.Append(runProps);
            run.Append(new Text("Привет"));

            Paragraph para = new Paragraph();
            para.Append(run);

            body.Append(para);
        }
    }
}