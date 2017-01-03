using System;

namespace VCardExtractPhotos {
    public static class Utils {
        public static string LF = System.Environment.NewLine;

        /// <summary>
        /// Error message.
        /// </summary>
        /// <param name="msg"></param>
        public static void errMsg(string msg) {
            Console.WriteLine("Error: " + msg);
        }

        /// <summary>
        /// Exception message.
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="ex"></param>
        public static void excMsg(string msg, Exception ex) {
            Console.WriteLine("Error: " + msg + LF + "Exception: " + ex + LF
            + ex.Message);
        }

        /// <summary>
        /// Warning message.
        /// </summary>
        /// <param name="msg"></param>
        public static void warnMsg(string msg) {
            Console.WriteLine("Warning: " + msg);
        }

        /// <summary>
        /// Information message.
        /// </summary>
        /// <param name="msg"></param>
        public static void infoMsg(string msg) {
            Console.WriteLine("Info: " + msg);
        }

    }
}
