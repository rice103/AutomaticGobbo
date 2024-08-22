using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows.Forms;

namespace MCIDEMO
{
    class MciPlayer
    {

        [DllImport("winmm.dll")]
        private static extern int mciSendString(String strCommand, StringBuilder strReturn, int iReturnLength, IntPtr hwndCallback);
        [DllImport("winmm.dll")]
        public static extern int mciGetErrorString(int errCode, StringBuilder errMsg, int buflen);
        [DllImport("winmm.dll")]
        public static extern int mciGetDeviceID(string lpszDevice);

        public const int WS_CHILD = 0x40000000;

        public MciPlayer()
        {

        }

        public MciPlayer(string filename, string alias)
        {
            FileInfo fi = new FileInfo(filename);
            String relocatedfile = Application.StartupPath + "\\unicodenamesupport\\" + alias + fi.Extension;
            if (File.Exists(relocatedfile))
                File.Delete(relocatedfile);
            if (!Directory.Exists("unicodenamesupport"))
                Directory.CreateDirectory("unicodenamesupport");
            File.Copy(filename, relocatedfile);

            _medialocation = filename;
            _alias = alias;
            LoadMediaFile(_medialocation, _alias);
        }

        int _deviceid = 0;

        public int Deviceid
        {
            get { return _deviceid; }
        }

        private bool _isloaded = false;

        public bool Isloaded
        {
            get { return _isloaded; }
            set { _isloaded = value; }
        }

        private string _medialocation = "";

        public string MediaLocation
        {
            get { return _medialocation; }
            set { _medialocation = value; }
        }
        private string _alias = "";

        public string Alias
        {
            get { return _alias; }
            set { _alias = value; }
        }


        public bool LoadMediaFile(string filename, string alias, IntPtr handle)
        {
            _medialocation = filename;
            _alias = alias;
            StopPlaying();
            CloseMediaFile();
            string Pcommand = "open mpegvideo!\"" + filename + "\" alias \"" + alias + "\" parent " + handle + " style " + WS_CHILD;
            int ret = mciSendString(Pcommand, null, 0, IntPtr.Zero);
            _isloaded = (ret == 0) ? true : false;
            if (_isloaded)
                _deviceid = mciGetDeviceID(_alias);
            else
            {
                StringBuilder sb = new StringBuilder(1024);
                int n = 1024;
                mciGetErrorString(ret, sb, n);
            }

            return _isloaded;

        }

        public bool LoadMediaFile(string filename, string alias)
        {
            _medialocation = filename;
            _alias = alias;
            StopPlaying();
            CloseMediaFile();
            string Pcommand = "open mpegvideo!\"" + filename + "\" alias \"" + alias + "\"";
            int ret = mciSendString(Pcommand, null, 0, IntPtr.Zero);
            _isloaded = (ret == 0) ? true : false;
            if (_isloaded)
                _deviceid = mciGetDeviceID(_alias);
            else
            {
                StringBuilder sb = new StringBuilder(1024);
                int n = 1024;
                mciGetErrorString(ret, sb, n);
            }

            return _isloaded;
        }

        public void PlayFromStart()
        {
            if (_isloaded)
            {
                string Pcommand = "play " + Alias + " from 0";
                int ret = mciSendString(Pcommand, null, 0, IntPtr.Zero);
            }
        }

        public void PlayFromStart(IntPtr callback)
        {
            if (_isloaded)
            {
                string Pcommand = "play " + Alias + " from 0 notify";
                int ret = mciSendString(Pcommand, null, 0, callback);
            }
        }


        public void PlayLoop()
        {
            if (_isloaded)
            {
                string Pcommand = "play " + Alias + " repeat";
                int ret = mciSendString(Pcommand, null, 0, IntPtr.Zero);
            }
        }

        public void CloseMediaFile()
        {
            string Pcommand = "close " + Alias;
            int ret = mciSendString(Pcommand, null, 0, IntPtr.Zero);
            _isloaded = false;

        }

        public void StopPlaying()
        {
            string Pcommand = "stop " + Alias;
            int ret = mciSendString(Pcommand, null, 0, IntPtr.Zero);
        }

        // Can't get this to work. May not be supported by the device
        //public void CaptureFrame(string destfile)
        //{
        //    string Pcommand = "capture " + Alias + " as \"" + destfile + "\"";
        //    int ret = mciSendString(Pcommand, null, 0, IntPtr.Zero);
        //    if (ret != 0)
        //    {
        //        StringBuilder sb = new StringBuilder(1024);
        //        int n = 1024;
        //        mciGetErrorString(ret, sb, n);
        //    }
        //}

        public void SetVideo(int x, int y, int width, int height)
        {
            string Pcommand = "put \"" + Alias + "\" window at " + x + " " + y + " " + width + " " + height;
            int ret = mciSendString(Pcommand, null, 0, IntPtr.Zero);
            if (ret != 0)
            {
                StringBuilder sb = new StringBuilder(1024);
                int n = 1024;
                mciGetErrorString(ret, sb, n);
            }

        }

        public string Status(string param)
        {
            string Pcommand = "status \"" + Alias + "\" " + param;
            StringBuilder sb = new StringBuilder(1024);
            int ret = mciSendString(Pcommand, sb, sb.Capacity, IntPtr.Zero);
            if (ret == 0)
                return sb.ToString();
            else
            {
                mciGetErrorString(ret, sb, sb.Capacity);
                return sb.ToString();
            }

        }

        public string MCISendString(string Pcommand)
        {
            StringBuilder sb = new StringBuilder(1024);
           // Pcommand = Pcommand.Replace("@@", Alias);
            int ret = mciSendString(Pcommand, sb, sb.Capacity, IntPtr.Zero);
            if (ret == 0)
                return sb.ToString();
            else
            {
                mciGetErrorString(ret, sb, sb.Capacity);
                return sb.ToString();
            }
        }

        public static void getAspectRatioSize(string source, ref int width, ref int height)
        {
            try
            {
                var s = source.Split(' ');
                float ar = float.Parse(s[2]) / float.Parse(s[3]);
                float c_ar = (float)width / (float)height;
                if (ar < c_ar) //height limiting
                {
                    width = (int)((float)height * (ar));
                }
                else   //width limiting
                {
                    height = (int)((float)width / (ar));
                }
            }
            catch { }

        }
    }
}
