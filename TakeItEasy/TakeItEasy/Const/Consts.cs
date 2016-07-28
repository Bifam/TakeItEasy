using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TakeItEasy.Const
{
    public static class Consts
    {
        /* File extension filter */
        public static string EXT_FILTER 
            = "Excel file (*.xlsx)|*.xlsx|Excel file (*.xls)|*.xls|CSV file (*.csv)|*.csv";
        /* Constant for DB accsess */
        public static string SERVER_STRING  = ".";
        public static string USER_STRING    = "sa";
        public static string PWD_STRING     = "tecse";

        /* Message title */
        // Information
        public static string MSG_INFO = "Information";
        // Error
        public static string MSG_ERR = "Error";
        // Warning
        public static string MSG_WARN = "Waring";

        /* Constant message */
        // Null path message
        public static string NEED_ENTER_PATH = "Please choose input folder and output file path!";
        // Path not exist
        public static string PATH_NOT_EXIST = "Inputed Folder path does not exist";
        // Generate file successfully
        public static string GEN_FILE_SUCCESS = "File was genetated successfully!";
        // Excel is not installed
        public static string EXCEL_NOT_INSTALL = "Excel is not properly installed!";
        // Field need filled
        public static string FIELD_NEED_FILLED = "Please fill all fields in Dialog!";

        /* Table name */
        // TB_DUTIES_MSG_MST
        public static string TB_DUTIES_MSG_MST = "TB_DUTIES_MSG_MST";
        // TB_RPT_HEAD_INFO
        public static string TB_RPT_HEAD_INFO = "TB_RPT_HEAD_INFO";

        // RPT_ID
        public static string RPT_ID = "RPT_ID";
        // HEAD_KIND
        public static string HEAD_KIND = "HEAD_KIND";
        // RPT_HEAD_SEQ_NO
        public static string RPT_HEAD_SEQ_NO = "RPT_HEAD_SEQ_NO";
        // RPT_HEAD_VAL
        public static string RPT_HEAD_VAL = "RPT_HEAD_VAL";
        // DUTIES_ID
        public static string DUTIES_ID = "DUTIES_ID";
        // DUTIES_MSG_ID
        public static string DUTIES_MSG_ID = "DUTIES_MSG_ID";
        // DUTIES_MSG
        public static string DUTIES_MSG = "DUTIES_MSG";

        public static string createConnString()
        {
            return "Server=" + SERVER_STRING + "; user Id = "+ USER_STRING + "; Password=" + PWD_STRING + ";";
        }
    }
}
