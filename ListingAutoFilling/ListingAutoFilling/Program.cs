using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ListingAutoFilling.Helpers.AddText;

namespace ListingAutoFilling
{
    public class Program
    {
        private const string DocFileName = @"C:\university\tst.docx";
        private const string PathToSearch = @"C:\university\danilka\RIS_CourseWork\ServerApp";
        
        public static void Main(string[] args)
        {
            using WordprocessingDocument wordDoc = WordprocessingDocument.Open(DocFileName, true);
            MainDocumentPart mainPart = wordDoc.MainDocumentPart ?? wordDoc.AddMainDocumentPart();
            
            Body body = mainPart.Document.AppendChild(new Body());

            var addTextToWordFileHelper = new AddTextToWordFileHelper();

            addTextToWordFileHelper.AddTextToWordFile(PathToSearch, body);
        }
    }
}