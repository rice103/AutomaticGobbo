/*
 * Created by SharpDevelop.
 * User: Rice Cipriani
 * Date: 07/05/2012
 * Time: 21:39
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using MidiExamples;
using System.Reflection;
using AutomaticGobbo;
using Sanford.Multimedia.Midi;
using System.Threading;
using System.Runtime.InteropServices;
using System.IO.Ports;
namespace GobboManuale
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
        String[] vectNote = { "C4", "CSharp4", "D4", "DSharp4", "E4", "F4", "FSharp4", "G4", "GSharp4", "A4", "ASharp4", "B4",
                            "C5", "CSharp5", "D5", "DSharp5", "E5", "F5", "FSharp5", "G5", "GSharp5", "A5", "ASharp5", "B5"};
        public List<long> listMessageRequested = new List<long>();
		bool played=false;
		bool txtLocked=false;
        int lastPeriod = 50;
        List<int> bookmarksIndexList = new List<int>();
		List<Line> lineList = new List<Line>();
        InputDevice inputDevice;
        OutputDevice outputDevice;
        private SynchronizationContext context;
        SerialPort comPort; 

		public MainForm(string[] args)
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
            int deviceId = -1;
			InitializeComponent();
            try
            {
                context = SynchronizationContext.Current;
                for (int i = 0; i < InputDevice.DeviceCount; i++)
                {
                    inputDevice = new InputDevice(i);
                    if (InputDevice.GetDeviceCapabilities(inputDevice.DeviceID).name == "CME U2MIDI")
                        break;
                }
                //inputDevice = new InputDevice(InputDevice.DeviceCount-1);
                
                if (inputDevice != null)
                {
                    //inputDevice.Open();
                    //inputDevice.StartReceiving(null);
                    inputDevice.ChannelMessageReceived += HandleChannelMessageReceived;
                    inputDevice.SysCommonMessageReceived += HandleSysCommonMessageReceived;
                    inputDevice.SysExMessageReceived += HandleSysExMessageReceived;
                    inputDevice.SysRealtimeMessageReceived += HandleSysRealtimeMessageReceived;
                    //inputDevice.Error += new EventHandler<ErrorEventArgs>(inDevice_Error);   
                    //inputDevice.Error += new EventHandler<ErrorEventArgs>(inDevice_Error);     
                    inputDevice.StartRecording();
                    Summarizer summarizer = new Summarizer(inputDevice, this);
                }
            }
            catch { }
            try
            {
                if (OutputDevice.DeviceCount>0)
                {
                    outputDevice = new OutputDevice(OutputDevice.DeviceCount - 1);
                }

            }
            catch { }

            try
            {
                //AudioTool.init();
            }
            catch { }
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
            
            String res = "C:\\Users\\Rice Cipriani\\Google Drive\\WWRY\\Testi.rtf"; 
            String res2 =  "C:\\Users\\Rice Cipriani\\Google Drive\\Scaletta + testi + intro_2.txt";
            if (args != null && args.Length > 0)
                res = args[0];
            if (args != null && args.Length > 1)
                res2 = args[1];
			if (res!=null && res!="")
			{
                if (res.EndsWith(".txt",true,System.Globalization.CultureInfo.CurrentCulture))
                    openOnText(res);
                if (res.EndsWith(".rtf", true, System.Globalization.CultureInfo.CurrentCulture))
                    openOnRtfText(res);
			}
            //button6.Tag = res;
            //button7.Tag = res2;

            String nameComPort = ""; 
            foreach (string s in SerialPort.GetPortNames())
                nameComPort = s;
            if (nameComPort != "")
            {
                comPort = new SerialPort(nameComPort);
                comPort.Open();
            }
		}

        public void nextLine()
        {
            ControlHelper.SuspendDrawing(textBox1);
            if (this.txtLocked && lineList.Count > 0)
            {
                int lineSel = lineSelected();
                if (lineSel < lineList.Count-1) 
                    textBox1.Select(lineList[lineSel + 1].getStart(), lineList[lineSel + 1].getText().Length);
            }
            doTheSelection();

            ControlHelper.ResumeDrawing(textBox1);
        }

        public void previousLine()
        {
            if (this.txtLocked && lineList.Count > 0)
            {
                int lineSel = lineSelected();
                if (lineSel>1)
                    textBox1.Select(lineList[lineSel - 1].getStart(), lineList[lineSel + 1].getText().Length);
            }
            doTheSelection();
        }

        public void loadFirstTime()
        {
            button6.PerformClick();
        }

        public void loadSecondTime()
        {
            button7.PerformClick();
        }

        public void play(int indexSong)
        {
            String[] listSong = { "C:\\Users\\Rice Cipriani\\Google Drive\\Belts2\\time_line_in_rock\\INTRO - Il nome della rosa.avi", 
                                    "C:\\Users\\Rice Cipriani\\Google Drive\\Belts2\\time_line_in_rock\\INTRO - Uno strano tipo di donna.avi",
                                    "C:\\Users\\Rice Cipriani\\Google Drive\\Belts2\\time_line_in_rock\\INTRO - Woodstock.avi" };
            
            string songFileName = listSong[indexSong];

            try
            {
                if (File.Exists(songFileName))
                {
                    System.Diagnostics.Process proc = new System.Diagnostics.Process();
                    proc.EnableRaisingEvents = false;
                    proc.StartInfo.FileName = songFileName;// "C:\\Program Files\\VideoLAN\\VLC\\vlc.exe";
                    proc.StartInfo.Arguments = "";// songFileName;
                    proc.Start();
                    //proc.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }

        }
        void openOnRtfText(string fileToOpen)
        {
            if (fileToOpen != null && fileToOpen != "")
            {
                try
                {
                    if (File.Exists(fileToOpen))
                    {
                        this.Text = "Automatic Gobbo - " + fileToOpen;//  + Path.GetFileName(fileToOpen);
                        StreamReader s = File.OpenText(fileToOpen);
                        textBox1.Rtf = s.ReadToEnd();
                        s.Close();
                        button7.PerformClick();
                        this.Tag = fileToOpen;
                    }
                }
                catch { }
            }
        }
        void openOnText(string fileToOpen)
        {
            if (fileToOpen != null && fileToOpen != "")
            {
                try
                {
                    if (File.Exists(fileToOpen))
                    {
                        this.Text = "Automatic Gobbo - " + fileToOpen;// Path.GetFileName(fileToOpen);
                        textBox1.Text = File.OpenText(fileToOpen).ReadToEnd();
                        textBox1.SelectAll();
                        textBox1.ForeColor = Color.White;
                        textBox1.Select(1, 0);
                        this.Tag = fileToOpen;
                    }
                }
                catch { }
            }

        }
		void Timer1Tick(object sender, EventArgs e)
		{
            //if (timer1.Tag !=null)
                timer1.Interval = lastPeriod;// int.Parse(timer1.Tag.ToString());
            timer1.Tag = null;
			if (!this.txtLocked){
				this.lockTextBoxAndLines();
			}
			int lineSel=lineSelected()+1;
			if (lineSel<textBox1.Lines.Length)
			{
				parseLine(lineSel);
			}
			else
			{
				timer1.Enabled=false;
			}
		}
		
		private void lockTextBoxAndLines(){
			lineList = new List<Line>();
            bookmarksIndexList = new List<int>();
			int tmpStart=0;
			for (int i=0;i<textBox1.Lines.Length;i++){
				string tmpLine=textBox1.Lines[i];
                if (tmpLine.Length>=3 && tmpLine.Substring(0, 3) == "[*]")
                    bookmarksIndexList.Add(i);
				lineList.Add(new Line(tmpStart,tmpLine, i));
				tmpStart+=tmpLine.Length+1;
			}
			txtLocked=true;
		}
		
		void Button2Click(object sender, EventArgs e)
		{
            //timer1.Interval = 50;
			played=true;
			if (!this.txtLocked){
				this.lockTextBoxAndLines();
			}
			int lineSel=lineSelected();
			parseLine(lineSel);

            timer1.Tag = timer1.Interval;
            timer1.Interval = 1;
            timer1.Enabled = true;
		}
		
		private void parseLine(int nLine){
			/* se è una riga semplice avvia il timer
			 * se è "[PAUSE]" ferma il timer
			 * se è "[T]300" modifica il tempo per riga
			 * se è "[EXT]shell" esegui con la shell
			 * ......
			 * */
			if (played && (nLine < lineList.Count))
			{
				timer1.Enabled=false;
				scrollText(nLine);

				textBox1.Select(lineList[nLine].getStart(),Math.Min(1,lineList[nLine].getText().Length));

				String command="";
				String param="";
				if (lineList[nLine].getText().Length>0){
					try{
					    command= lineList[nLine].getText().Split('[')[1].Split(']')[0];
					    param=lineList[nLine].getText().Split('[')[1].Split(']')[1];
					}
					catch{}
				}
				switch (command.ToLower()){
                    case "vlc":
                        {
                            try
                            {
                                //System.Diagnostics.Process.Start(param);
                                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                                proc.EnableRaisingEvents=false;
                                proc.StartInfo.FileName = "C:\\Program Files\\VideoLAN\\VLC\\vlc.exe";
                                proc.StartInfo.Arguments = param.Replace("'","\u0022");
                                //proc.StartInfo.FileName=;
                                proc.Start();
                            }
                            catch { }
                            timer1.Enabled = true;
                            break;
                        }
                    case "com":
                        {
                            try
                            {
                                String strPort = param.TrimEnd();
                                int port = int.Parse(strPort.Replace("[\\D]", ""));
                                if (comPort != null)
                                    comPort.Close();
                                comPort = new SerialPort("COM" + port);
                                comPort.Open();
                            }
                            catch { }
                            timer1.Enabled = true;
                            break;
                        }
                    case "dmx":
                        {
                            if (comPort != null)
                            {
                                try
                                {
                                    comPort.Write(param.TrimEnd());
                                }
                                catch { }
                            }
                            timer1.Enabled = true;
                            break;
                        }
                    case "prg":
                        {
                            String strPrgNumber = param.TrimEnd();
                            int prgNumber = int.Parse(strPrgNumber.Replace("[\\D]", ""));
                            if (prgNumber>=0 && prgNumber<128)
                                outputDevice.Send(new ChannelMessage(ChannelCommand.ProgramChange, 0, prgNumber));
                            timer1.Enabled = true;
                            break;
                        }
                    case "mid":
                        {
                            try
                            {
                                String tonica = param.TrimEnd().Replace("-","").Replace("7","").ToUpper().Replace("#", "Sharp") + "4";
                                int inVect = new List<String>(vectNote).IndexOf(tonica);
                                String terza = vectNote[inVect + 4];
                                String quinta = vectNote[inVect + 7];
                                if (param.Contains("-"))
                                {
                                    terza = vectNote[inVect + 3];
                                }
                                outputDevice.Send(new ChannelMessage(ChannelCommand.NoteOn, 0, inVect, 127));//  SendNoteOn(Channel.Channel1, (Pitch)Enum.Parse(typeof(Pitch), tonica), 127);
                                outputDevice.Send(new ChannelMessage(ChannelCommand.NoteOn, 0, inVect + 4, 127));
                                outputDevice.Send(new ChannelMessage(ChannelCommand.NoteOn, 0, inVect + 7, 127)); 
                                if (param.Contains("7"))
                                    outputDevice.Send(new ChannelMessage(ChannelCommand.NoteOn, 0, inVect + 10, 127)); //outputDevice.SendNoteOn(Channel.Channel1, (Pitch)Enum.Parse(typeof(Pitch), vectNote[inVect + 10]), 127);
                                
                                Application.DoEvents();
                                outputDevice.Send(new ChannelMessage(ChannelCommand.NoteOff, 0, inVect , 127)); //outputDevice.SendNoteOff(Channel.Channel1, (Pitch)Enum.Parse(typeof(Pitch), inVect), 127);
                                outputDevice.Send(new ChannelMessage(ChannelCommand.NoteOff, 0, inVect + 4, 127)); //outputDevice.SendNoteOff(Channel.Channel1, (Pitch)Enum.Parse(typeof(Pitch), inVect), 127);
                                outputDevice.Send(new ChannelMessage(ChannelCommand.NoteOff, 0, inVect + 5, 127)); //outputDevice.SendNoteOff(Channel.Channel1, (Pitch)Enum.Parse(typeof(Pitch), inVect), 127);
                                outputDevice.Send(new ChannelMessage(ChannelCommand.NoteOff, 0, inVect + 10, 127)); //outputDevice.SendNoteOff(Channel.Channel1, (Pitch)Enum.Parse(typeof(Pitch), inVect), 127);
                                //outputDevice.SendNoteOff(Channel.Channel1, (Pitch)Enum.Parse(typeof(Pitch), inVect + 4), 127);
                                //outputDevice.SendNoteOff(Channel.Channel1, (Pitch)Enum.Parse(typeof(Pitch), inVect+7), 127);
                                //outputDevice.SendNoteOff(Channel.Channel1, (Pitch)Enum.Parse(typeof(Pitch), inVect + 10), 127);
                            }
                            catch { }
                            timer1.Tag = timer1.Interval;
                            timer1.Interval = 1;
                            timer1.Enabled = true;
                        }
                        break;
                    case "pause":
					{
						//do nothing
						break;
					}
                    case "audioload":
                    {
                        try
                        {
                            AudioTool.SoundLoad(param.Split(',')[0], int.Parse(param.Split(',')[1]));
                        }
                        catch { }
                        break;
                    }
                    case "audioplay":
                    {
                        AudioTool.play(int.Parse(param));
                        break;
                    }
                    case "audiostop":
                    {
                        AudioTool.stopAll();
                        break;
                    }
					case "ext":
					{
						try{
							System.Diagnostics.Process.Start(param);
							/*System.Diagnostics.Process proc = new System.Diagnostics.Process();
							proc.EnableRaisingEvents=false;
							proc.StartInfo.FileName=param;
							//proc.StartInfo.FileName=;
							proc.Start();*/
						}
						catch{}
						timer1.Enabled=true;
						break;
					}
                    case "*":
                    {
                        try
                        {
                            lastPeriod = 50;
                            timer1.Interval = 50;
                            timer1.Enabled = true;
                        }
                        catch { }
                        break;
                    }
                    case "/":
                    case "/2":
                    {
                        timer1.Tag = timer1.Interval;
                        timer1.Interval = int.Parse(Math.Ceiling(((double)lastPeriod) / (2.0)).ToString());
                        timer1.Enabled = true;
                        break;
                    }
                    case "/4":
                    {
                        timer1.Tag = timer1.Interval;
                        timer1.Interval = int.Parse(Math.Ceiling(((double)lastPeriod) / (4.0)).ToString());
                        timer1.Enabled = true;
                        break;
                    }
                    case "*2":
                    {
                        timer1.Tag = timer1.Interval;
                        timer1.Interval = int.Parse(Math.Ceiling(((double)lastPeriod) * (2.0)).ToString());
                        timer1.Enabled = true;
                        break;
                    }
                    case "-":
                    {
                        timer1.Tag = timer1.Interval;
                        timer1.Interval = 1;
                        timer1.Enabled = true;
                        break;
                    }
                    case "t":
					{
                        try
                        {
                            lastPeriod = Int16.Parse(param);
                            timer1.Interval = lastPeriod;
                            timer1.Enabled = true;
                        }
                        catch { }
                        break;
					}
                    case "t+pause":
                    {
                        try
                        {
                            lastPeriod = Int16.Parse(param);
                            timer1.Interval = lastPeriod;
                            timer1.Enabled = false;
                        }
                        catch { }
                        break;
                    }
                    default:
                    {
                        timer1.Enabled = true;
                        break;
                    }
				}
			}
		}

        private void scrollText(int nLine)
        {
            scrollText(nLine, true);
        }

        private void scrollText(int nLine, bool dirDown)
        {
            textBox1.Focus();
            if (nLine >3)
            {
                textBox1.Select(lineList[nLine - 3].getStart(), 0);
                textBox1.ScrollToCaret();
            }
        }
		
		private int lineSelected(){
			for (int i=0;i<lineList.Count;i++){
				if ((textBox1.SelectionStart>=lineList[i].getStart()) && (textBox1.SelectionStart<(lineList[i].getText().Length + lineList[i].getStart() +2 ))){
					return i+1;
				}
			}
			return 0;
		}
		
		
		void TextBox1TextChanged(object sender, EventArgs e)
		{
            if (txtLocked)
            {
                textBox1.Undo();
                txtLocked = false;
                played = false;
                timer1.Enabled = false;
            }
		}
		
		void TextBox1Click(object sender, EventArgs e)
		{
            try
            {
                if (this.txtLocked && lineList.Count > 0)
                {
                    int lineSel = lineSelected()-1;
                    parseLine(lineSel);
                }
            }
            catch { }
		}
        void doTheSelection()
        {
            try
            {
                if (this.txtLocked && lineList.Count > 0)
                {
                    int lineSel = lineSelected();
                    parseLine(lineSel);
                }
            }
            catch { }

        }

		void TextBox1KeyDown(object sender, KeyEventArgs e)
		{
            if (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right)
            {
                txtLocked = false;
                played = false;
                timer1.Enabled = false;
            }
            if (e.KeyCode == Keys.Down)
            {
                nextLine();
                if (this.txtLocked)
                    e.SuppressKeyPress = true;
            }
            if (e.KeyCode == Keys.Up)
            {
                previousLine(); 
                if (this.txtLocked)
                    e.SuppressKeyPress = true;
            }
		}
		
		void Button1Click(object sender, EventArgs e)
		{
            txtLocked = false;
            played = false;
            timer1.Enabled = false;
			textBox1.Focus();
		}
		
		void Button3Click(object sender, EventArgs e)
		{
			try{
				System.Diagnostics.Process proc = new System.Diagnostics.Process();
				proc.EnableRaisingEvents=false;
				proc.StartInfo.FileName=Environment.GetFolderPath(Environment.SpecialFolder.System) + "\\osk.exe";
				proc.Start();
			}
			catch{}
			
		}
		
		void Button4Click(object sender, EventArgs e)
		{
			openFileDialog1.ShowDialog();
			String res= openFileDialog1.FileName;
			if (res!=null && res!="")
			{
				try{
					if  (File.Exists(res))
					{
                        if (res.EndsWith(".txt", true, System.Globalization.CultureInfo.CurrentCulture))
                            openOnText(res);
                        if (res.EndsWith(".rtf", true, System.Globalization.CultureInfo.CurrentCulture))
                            openOnRtfText(res);
				     	this.Tag=res;
					}
				}
				catch{}
			}
			
		}
		
		void MainFormLoad(object sender, EventArgs e)
		{
            
            //textBox1.Font= new Font(textBox1.Font.Name,(int)(Screen.PrimaryScreen.Bounds.Height/15));
		}
		
		void Button5Click(object sender, EventArgs e)
		{
			if (this.Tag!=null)
			{
			saveFileDialog1.FileName=this.Tag.ToString();
			saveFileDialog1.InitialDirectory=this.Tag.ToString();
			}
			saveFileDialog1.ShowDialog();
			string res= saveFileDialog1.FileName;
			if (res!=null && res!="")
			{
				try{
					if  (File.Exists(res))
						File.Delete(res);
					File.WriteAllText (res,textBox1.Rtf);
					this.Tag=res;
				}
				catch{}
			}
		}
		
		void SaveFileDialog1FileOk(object sender, System.ComponentModel.CancelEventArgs e)
		{
			
		}

        private void button6_Click(object sender, EventArgs e)
        {
            //if (MessageBox.Show("Vuoi salvare prima?", "", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            //    button5.PerformClick();
            //openOnText(button6.Tag.ToString());
            //button1.PerformClick();
            //button2.PerformClick();
            //textBox1.SelectAll();
            textBox1.BackColor = Color.White;
            //textBox1.DeselectAll();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //if (MessageBox.Show("Vuoi salvare prima?", "", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            //    button5.PerformClick(); 
            //openOnText(button7.Tag.ToString());
            //button1.PerformClick();
            //button2.PerformClick();
            //textBox1.SelectAll();
            textBox1.BackColor = Color.Black;
            //textBox1.DeselectAll();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                inputDevice.StopRecording();
                inputDevice.Close();
                
            }
            catch { }
            try
            {
                if (comPort != null && comPort.IsOpen)
                    comPort.Close();
                //inputDevice.RemoveAllEventHandlers();
            }
            catch { }
            if (MessageBox.Show("Vuoi salvare prima?", "", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                button5.PerformClick();
            Environment.Exit(0);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            play(0);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            play(1);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            play(2);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            play(3);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            play(4);
        }

        private void inDevice_Error(object sender, ErrorEventArgs e)
        {
            MessageBox.Show(e.ToString(), "Error!",
                   MessageBoxButtons.OK, MessageBoxIcon.Stop);
        }

        private void HandleChannelMessageReceived(object sender, ChannelMessageEventArgs e)
        {
            //context.Post(delegate(object dummy)
            //{
            //    //channelListBox.Items.Add(
            //    //    e.Message.Command.ToString() + '\t' + '\t' +
            //    //    e.Message.MidiChannel.ToString() + '\t' +
            //    //    e.Message.Data1.ToString() + '\t' +
            //    //    e.Message.Data2.ToString());

            //    //channelListBox.SelectedIndex = channelListBox.Items.Count - 1;
            //}, null);
            if (!listMessageRequested.Contains(e.Message.Message))
                listMessageRequested.Add(e.Message.Message);
        }

        private void HandleSysExMessageReceived(object sender, SysExMessageEventArgs e)
        {
            //context.Post(delegate(object dummy)
            //{
            //    //string result = "\n\n"; ;

            //    //foreach (byte b in e.Message)
            //    //{
            //    //    result += string.Format("{0:X2} ", b);
            //    //}

            //    //sysExRichTextBox.Text += result;
            //}, null);
        }

        private void HandleSysCommonMessageReceived(object sender, SysCommonMessageEventArgs e)
        {
            //context.Post(delegate(object dummy)
            //{
            //    //sysCommonListBox.Items.Add(
            //    //    e.Message.SysCommonType.ToString() + '\t' + '\t' +
            //    //    e.Message.Data1.ToString() + '\t' +
            //    //    e.Message.Data2.ToString());

            //    //sysCommonListBox.SelectedIndex = sysCommonListBox.Items.Count - 1;
            //}, null);
        }

        private void HandleSysRealtimeMessageReceived(object sender, SysRealtimeMessageEventArgs e)
        {
            //context.Post(delegate(object dummy)
            //{
            //    //sysRealtimeListBox.Items.Add(
            //    //    e.Message.SysRealtimeType.ToString());

            //    //sysRealtimeListBox.SelectedIndex = sysRealtimeListBox.Items.Count - 1;
            //}, null);
            listMessageRequested.Add(e.Message.Message);
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            while (this.listMessageRequested.Count>0)
            {
                //Midi.Instrument msg = listMessageRequested[0];
                long msg = listMessageRequested[0];
                
                switch (msg)
                {
                    //case 192://Preset 1  :preset k = ((n - 192) / 256) + 1
                    //    {
                    //        break;
                    //    }
                    //case 448://Preset 2  :preset 2 = ((448 - 192) / 256) +1
                    //    {
                    //        break;
                    //    }
                    case 4222896://Double
                        {
                            break;
                        }
                    case 28336://Harmony
                        {
                            break;
                        }
                    case 28848://Reverber
                        {
                            break;
                        }
                    case 4221104://FX
                        {
                            break;
                        }
                    case 4224432://Delay
                        {
                            break;
                        }
                    case 4224176://nMOD
                        {
                            break;
                        }

                    case 448://16://Instrument.DrawbarOrgan:
                        {
                            zoomplus();
                            break;
                        }
                    case 2752://15://Instrument.Dulcimer:
                        {
                            zoomminus();
                            break;
                        }
                    case 4288://17://Instrument.PercussiveOrgan:
                        {
                            togglefullscreen();
                            break;
                        }
                    case 3264://12://(Instrument.Marimba):
                        {
                            button2.PerformClick();
                            break;
                        }
                    case 3776://14://(Instrument.TubularBells):
                        {
                            nextLine();
                            break;
                        }
                    case 3520://13://(Instrument.Xylophone):
                        {
                            previousLine();
                            break;
                        }
                    case 10://(Instrument.MusicBox):
                        {
                            loadFirstTime();
                            break;
                        }
                    case 1://(Instrument.BrightAcousticPiano):
                        {
                            loadSecondTime();
                            break;
                        }
                    case 4800://18://(Instrument.RockOrgan):
                        {
                            button13.PerformClick();
                            break;
                        }
                    case 5056://19://(Instrument.ChurchOrgan):
                        {
                            button14.PerformClick();
                            break;
                        }
                }
                listMessageRequested.Remove(msg);
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            int lastBookmarks = 0;
            for (int i = 0; i < bookmarksIndexList.Count; i++) 
            {
                if (lineSelected() >= bookmarksIndexList[i])
                    lastBookmarks = i;
            }
            if (lastBookmarks>=bookmarksIndexList.Count-1)
                lastBookmarks--;
            if (lastBookmarks>=0)
                textBox1.Select(lineList[bookmarksIndexList[lastBookmarks + 1]].getStart(), 0);
            else
                textBox1.Select(0, 0);
            doTheSelection();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            int lastBookmarks = 0;
            for (int i = 0; i < bookmarksIndexList.Count; i++)
            {
                if (lineSelected() > bookmarksIndexList[i])
                    lastBookmarks = i;
            }
            lastBookmarks--;
            if (lastBookmarks >= 0 && bookmarksIndexList.Count>0)
            {
                if (bookmarksIndexList[lastBookmarks] > lineList.Count)
                    lastBookmarks--;
                textBox1.Select(lineList[bookmarksIndexList[lastBookmarks]].getStart(), 0);
            }
            else
                textBox1.Select(0, 0);
            doTheSelection();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            textBox1.SelectionFont = new Font("Arial Black", 36, FontStyle.Bold);
            textBox1.SelectionColor = button15.BackColor;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            textBox1.SelectionFont = new Font("Arial Black", 36, FontStyle.Bold);
            textBox1.SelectionColor = button16.BackColor;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            textBox1.SelectionFont = new Font("Arial Black", 36, FontStyle.Bold);
            textBox1.SelectionColor = button17.BackColor;
        }

        private void button18_Click_1(object sender, EventArgs e)
        {
            textBox1.SelectionFont = new Font("Arial Black", 36, FontStyle.Bold);
            textBox1.SelectionColor = button18.BackColor;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            textBox1.SelectionFont = new Font("Arial Black", 36, FontStyle.Bold);
            textBox1.SelectionColor = button19.BackColor;
        }

        private void button21_Click(object sender, EventArgs e)
        {
            textBox1.SelectionFont = new Font("Arial Black", 36, FontStyle.Bold);
            textBox1.SelectionColor = button21.BackColor;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            textBox1.SelectionFont = new Font("Arial Black", 36, FontStyle.Bold);
        }

        private void button20_Click(object sender, EventArgs e)
        {
            textBox1.SelectionFont = new Font("Arial Black", 8, FontStyle.Regular);
        }

        private void zoomplus()
        {
            zoom = Math.Min(10.0F, zoom + 0.2F);
            textBox1.ZoomFactor = zoom;
        }

        private void zoomminus()
        {
            zoom = Math.Max(0.4F, zoom - 0.2F);
            textBox1.ZoomFactor = zoom;
        }

        private void togglefullscreen()
        {
            if (tableLayoutPanel1.ColumnStyles[1].Width == 100)
            {
                tableLayoutPanel1.ColumnStyles[1].Width = 0;
                FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                textBox1.ScrollBars = RichTextBoxScrollBars.None;
            }
            else if (tableLayoutPanel1.ColumnStyles[1].Width == 0)
            {
                tableLayoutPanel1.ColumnStyles[1].Width = 100;
                FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                textBox1.ScrollBars = RichTextBoxScrollBars.ForcedVertical;
            }
        }

        float zoom = 1;
        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F11)
            {
                togglefullscreen();
            }
            if (e.Control && e.Shift)
            {
                if (e.KeyCode == Keys.OemMinus)
                {
                    zoomminus();
                }
                else if (e.KeyCode == Keys.Oemplus)
                {
                    zoomplus();
                }
            }
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            timer3.Enabled = false;
            timer3.Tag = null;
        }

        private void button23_Click(object sender, EventArgs e)
        {
            //Int16 pippo;
            //pippo = (int)Midi.Instrument.DrawbarOrgan;
            //pippo = (int)Midi.Instrument.Dulcimer;
            //pippo = (int)Midi.Instrument.PercussiveOrgan;
            //pippo = (int)Midi.Instrument.Marimba;
            //pippo = (int)Midi.Instrument.TubularBells;
            //pippo = (int)Midi.Instrument.Xylophone;
            //pippo = (int)Midi.Instrument.MusicBox;
            //pippo = (int)Midi.Instrument.BrightAcousticPiano;
            //pippo = (int)Midi.Instrument.RockOrgan;
            //pippo = (int)Midi.Instrument.ChurchOrgan;
            //int index = OutputDevice.InstalledDevices.Count-1;
            //OutputDevice.InstalledDevices[index].Open();
            //OutputDevice.InstalledDevices[index].SendProgramChange(Channel.Channel1, Instrument.Marimba);
            //OutputDevice.InstalledDevices[index].SendControlChange(Channel.Channel1, Midi.Control.Volume, 127); //volume output 0=mute 127=0db
            //OutputDevice.InstalledDevices[index].Close();
            outputDevice.Send( new ChannelMessage(ChannelCommand.ProgramChange,0,13));
        }

	}
    public class Summarizer
    {
        MainForm callerForm;
        public Summarizer(InputDevice inputDevice, MainForm callerForm)
        {
            this.callerForm = callerForm;
            this.inputDevice = inputDevice;
            //pitchesPressed = new Dictionary<Pitch, bool>();
            //inputDevice.NoteOn += new InputDevice.NoteOnHandler(this.NoteOn);
            //inputDevice.NoteOff += new InputDevice.NoteOffHandler(this.NoteOff);
            //inputDevice.ControlChange += new InputDevice.ControlChangeHandler(inputDevice_ControlChange);
            //inputDevice.ProgramChange += new InputDevice.ProgramChangeHandler(inputDevice_ProgramChange);
        }
        //public void inputDevice_ControlChange(Midi.ControlChangeMessage msg)
        //{
        //    lock (this)
        //    {
        //        //callerForm.listMessageRequested.Add(msg);
        //    }
        //}

        //public void inputDevice_ProgramChange(Midi.ProgramChangeMessage msg)
        //{
        //    lock (this)
        //    {
        //        if (callerForm.timer3.Tag == null)
        //        {
        //            callerForm.listMessageRequested.Add(msg.Instrument);
        //            callerForm.timer3.Tag = "stopProgramChange";
        //            callerForm.timer3.Enabled = true;
        //        }
        //    }
        //}

        //public void NoteOn(NoteOnMessage msg)
        //{
        //    lock (this)
        //    {
        //        pitchesPressed[msg.Pitch] = true;
        //    }
        //}

        //public void NoteOff(NoteOffMessage msg)
        //{
        //    lock (this)
        //    {
        //        pitchesPressed.Remove(msg.Pitch);
        //    }
        //}

        private InputDevice inputDevice;
        //private Dictionary<Pitch, bool> pitchesPressed;

       
    }
    public static class ControlHelper
    {
        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hwnd, int wMsg, int wParam, int lParam);

        private const int WM_SETREDRAW = 0xB;

        public static void SuspendDrawing(Control target)
        {
            SendMessage(target.Handle, WM_SETREDRAW, 0, 0);
        }

        public static void ResumeDrawing(Control target)
        {
            SendMessage(target.Handle, WM_SETREDRAW, 1, 0);
            target.Refresh();
        }
    } 
}
