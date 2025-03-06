using ListingAutoFilling.Helpers.GetText.Data;

namespace ListingAutoFilling.Helpers.GetText;

public interface IGetCodeTextHelper
{
    FileData GetFileData(string path, string fileName);
}