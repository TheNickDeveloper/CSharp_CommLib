using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace VstoHelperTest.Helper
{
    public static class CommonFunctionHelper
    {
        //----------------------------------------------
        //Dictionary generation
        //----------------------------------------------

        public static Dictionary<dynamic, dynamic> GenerateDictionary(dynamic[,] arrTarget
            , int keyIndex, int valueIndex)
        {
            var dicTarget = new Dictionary<dynamic, dynamic>();

            for (var i = 0; i < arrTarget.GetLength(0); i++)
                if (dicTarget.ContainsKey(arrTarget[i, keyIndex]) == false)
                    dicTarget.Add(arrTarget[i, keyIndex], arrTarget[i, valueIndex]);

            return dicTarget;
        }

        //----------------------------------------------
        // Value contents handling
        //----------------------------------------------

        public static string Left(string target, int contentLength)
        {
            return target.Substring(0, contentLength);
        }

        public static string Right(string target, int contentLength)
        {
            return target.Substring(target.Length - contentLength, contentLength);
        }

        public static string Mid(string target, int stratIndex, int contentLength)
        {
            return target.Substring(stratIndex, contentLength);
        }

        //----------------------------------------------
        // File Browse
        //----------------------------------------------

        public static string GetFilePathByBrowse()
        {
            var openfileDia = new OpenFileDialog();
            return openfileDia.ShowDialog() == DialogResult.OK 
                ? openfileDia.FileName 
                : "Action has been cancled.";
        }

        public static string GetFolderPathByBrowse()
        {
            var openFolderDia = new FolderBrowserDialog();
            return openFolderDia.ShowDialog() == DialogResult.OK 
                ? openFolderDia.SelectedPath 
                : "Action has been cancled.";
        }

        //----------------------------------------------
        // Error handling
        //----------------------------------------------

        public static void ErrorHandling(string errorMessage)
        {
            var certifyException = new Exception(errorMessage);
            throw certifyException;
        }
    }
}
