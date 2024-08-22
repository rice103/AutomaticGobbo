using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NAudio.Wave;
using System.IO;
using System.Windows.Forms;

namespace AutomaticGobbo
{
    public class AudioTool
    {
        public static IWavePlayer waveOutDevice;
        public static List<WaveStream> wBuf;
        public static void Mp3Converter(string mp3file, string wavfile)
        {
            using (var reader = new Mp3FileReader(mp3file))
            {
                using (var writer = new WaveFileWriter(wavfile, new WaveFormat()))
                {
                    var buf = new byte[4096];
                    for (; ; )
                    {
                        var cnt = reader.Read(buf, 0, buf.Length);
                        if (cnt == 0) break;
                        writer.WriteData(buf, 0, cnt);
                    }
                }
            }
        }
        public static bool SoundLoad(String file, int buffNumber)
        {
            string dest = Environment.CurrentDirectory + "\\" + buffNumber + ".wav";
            if (File.Exists(dest))
                File.Delete(dest);
            if (file.ToLower().EndsWith(".mp3"))
            {
                Mp3Converter(file,dest);
            }
            else if (file.ToLower().EndsWith(".wav"))
            {
                File.Copy(file, dest);
            }
            if (File.Exists(dest))
            {
                wBuf[buffNumber] = CreateInputStream(dest);
                return true;
            }
            return false; 
        }
        public static void init()
        {
            waveOutDevice = new AsioOut();
            wBuf = (new WaveStream[10]).ToList<WaveStream>();
        }
        public static WaveStream CreateInputStream(string fileName)
        {
            WaveChannel32 inputStream;
            if (fileName.EndsWith(".wav"))
            {
                WaveStream readerStream = new WaveFileReader(fileName);
                if (readerStream.WaveFormat.Encoding != WaveFormatEncoding.Pcm)
                {
                    readerStream = WaveFormatConversionStream.CreatePcmStream(readerStream);
                    readerStream = new BlockAlignReductionStream(readerStream);
                }
                if (readerStream.WaveFormat.BitsPerSample != 16)
                {
                    var format = new WaveFormat(readerStream.WaveFormat.SampleRate,
                       16, readerStream.WaveFormat.Channels);
                    readerStream = new WaveFormatConversionStream(format, readerStream);
                }
                inputStream = new WaveChannel32(readerStream);
            }
            else
            {
                throw new InvalidOperationException("Unsupported extension");
            }
            return inputStream;
        }
        public static void play(int bufNumber)
        {
            stopAll();
            
            try
            {
                waveOutDevice.Init(wBuf[bufNumber]);
            }
            catch { }
            waveOutDevice.Play();
        }
        public static void stopAll()
        {
            try
            {
                waveOutDevice.Stop();
            }
            catch { }
            try
            {
                waveOutDevice.Dispose();
                waveOutDevice = new AsioOut();
            }
            catch { }
        }
    }
    
}