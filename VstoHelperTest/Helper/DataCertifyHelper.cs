using System.IO;

namespace VstoHelperTest.Helper
{
    public static class DataCertifyHelper
    {
        public static void VerifyIsEmptyContents(string target, string subject)
        {
            if (string.IsNullOrEmpty(target))
                CommonFunctionHelper.ErrorHandling(subject + " cannot be blank.");
        }

        public static void VerifyFilePath(string path, string subject)
        {
            VerifyIsEmptyContents(path,subject);

            if (File.Exists(path) == false)
                CommonFunctionHelper.ErrorHandling(subject + " did not exsit.");
        }

        public static void VerifyFolderPath(string path, string subject)
        {
            VerifyIsEmptyContents(path, subject);

            if (Directory.Exists(path) == false)
                CommonFunctionHelper.ErrorHandling(subject + " did not exsit.");
        }
    }
}
