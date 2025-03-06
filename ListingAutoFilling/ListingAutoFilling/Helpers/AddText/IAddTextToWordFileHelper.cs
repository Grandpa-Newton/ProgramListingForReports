using DocumentFormat.OpenXml.Wordprocessing;

namespace ListingAutoFilling.Helpers.AddText;

public interface IAddTextToWordFileHelper
{
    void AddTextToWordFile(string folderToSearch, Body body);
}