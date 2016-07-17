using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SpeechLib;
using System.Timers;
using System.IO;
using System.Media;

//using DotNetSpeech;


using System.Speech.AudioFormat;
using System.Speech.Recognition;

namespace MusicPlayer
{
    public partial class SpeechRecognition : Form
    {
        private SpeechLib.ISpeechRecoContext wavRecoContext = null;
        private SpeechLib.SpFileStream InputWAV = null;
        private SpeechLib.ISpeechRecoGrammar Grammar = null;
        private String _WAVFile = null;
        private string strData = "No recording yet";
        private String allRecognized = "";

        //编辑歌词时，歌词文件路径
        private string lyricpath;

        public SpeechRecognition()
        {
            InitializeComponent();
        }

        // Create a simple handler for the SpeechRecognized event.
        //void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        //{
        //    MessageBox.Show("Speech recognized: " + e.Result.Text);
        //    textBoxX1.Text += e.Result.Text;
        //}

        //private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    /*FileStream stream = File.Open(lstSongs.Text, FileMode.Open);
        //    SpeechAudioFormatInfo speechaudioformatinfo=new SpeechAudioFormatInfo(100,
        //    System.Speech.AudioFormat.AudioBitsPerSample.Eight,
        //    System.Speech.AudioFormat.AudioChannel.Mono);//单声道的通道
        //    //	System.Speech.AudioFormat.AudioChannel.Stereo  立体声通道

        //    SpeechRecognitionEngine speechRecognitionEngine = new SpeechRecognitionEngine();
        //    Choices colors = new Choices();
        //    colors.Add(new string[] {"海","天", "星星", "点灯", "孩子" });

        //    // Create a GrammarBuilder object and append the Choices object.
        //    GrammarBuilder gb = new GrammarBuilder();
        //    gb.Append(colors);

        //    // Create the Grammar instance and load it into the speech recognition engine.
        //    Grammar g = new Grammar(gb);
        //    speechRecognitionEngine.LoadGrammar(g);

        //    speechRecognitionEngine.SetInputToAudioStream(stream, speechaudioformatinfo);

        //    speechRecognitionEngine.SpeechRecognized +=
        //      new EventHandler<SpeechRecognizedEventArgs>(sre_SpeechRecognized);
            
        //    stream.Close();
        //     */  
        //}
        
        private void wavRecoContext_Recognition(int StreamNumber, object StreamPosition, SpeechRecognitionType RecognitionType, ISpeechRecoResult Result)
        {
            strData = Result.PhraseInfo.GetText(0, -1, true);
            allRecognized += textBoxX1.Text;

            textBoxX1.Text = strData;

            string path = _WAVFile.Substring(0, _WAVFile.Length - 4);
            path += ".lrc";
            FileStream fs = new FileStream(path, FileMode.Append);
            StreamWriter writer = new StreamWriter(fs);
            writer.WriteLine();

            //把原来的转换为byte数组，而且是utf_8编码的
            //Byte[] change = System.Text.Encoding.UTF8.GetBytes(allRecognized.ToCharArray());
            //用convet直接转换
            //Byte[] changde = Encoding.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, change);
            //allRecognized=changde.ToString();

            //allRecognized = Encoding.ASCII.GetString(Encoding.ASCII.GetBytes(allRecognized.ToCharArray()));//其中ss为你的utf8的数据
            writer.Write(allRecognized, Encoding.ASCII);
            writer.Flush();
            fs.Close();
            allRecognized = "";


            //Encoding ecp1252 = Encoding.GetEncoding(1252);
            //StreamReader sr = new StreamReader(path, Encoding.Unicode, false);
            //歌词文件编码转化
            //String newPath = path.Substring(0, _WAVFile.Length - 4) + ".lrc";
            //StreamWriter sw = new StreamWriter(newPath, false, ecp1252);
            //sw.Write(sr.ReadToEnd());
            //sw.Close();
            //sr.Close();
        }
        private void wavRecoContext_EndStream(int StreamNumber, object StreamPosition, bool f)
        {

            // ' Disable dictation

            Grammar.DictationSetState(SpeechRuleState.SGDSInactive);

            string path = _WAVFile.Substring(0, _WAVFile.Length - 4);
            path += ".lrc";
            FileStream fs = new FileStream(path, FileMode.Create);
            StreamWriter writer = new StreamWriter(fs);
            writer.Write(allRecognized);
            writer.Flush();
            fs.Close();
            allRecognized = "";
        }
        private void labelX1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();

                dialog.Title = "Select a Speech file";

                dialog.ShowDialog();

                _WAVFile = dialog.FileName;

                if (_WAVFile == null) return;

                comboBoxEx1.Items.Add(dialog.FileName);
                comboBoxEx1.SelectedText = dialog.FileName;

                wavRecoContext = new SpeechLib.SpInProcRecoContext();

                ((SpInProcRecoContext)wavRecoContext).Recognition +=

                new _ISpeechRecoContextEvents_RecognitionEventHandler(wavRecoContext_Recognition);

                ((SpInProcRecoContext)wavRecoContext).EndStream +=
                    new _ISpeechRecoContextEvents_EndStreamEventHandler(wavRecoContext_EndStream);

                Grammar = wavRecoContext.CreateGrammar(0);

                Grammar.DictationLoad("", SpeechLoadOption.SLOStatic);

                InputWAV = new SpFileStream();

                InputWAV.Open(@_WAVFile, SpeechStreamFileMode.SSFMOpenForRead, false);

                wavRecoContext.Recognizer.AudioInputStream = InputWAV;

                Grammar.DictationSetState(SpeechRuleState.SGDSActive);
            }
            catch (Exception er)
            {
                //MessageBox.Show("An Error Occured!", "SpeechApp", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Console.WriteLine(er.ToString());
            }
            
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            try
            {
                /*SpVoice voice = new SpVoice();
                voice.Voice = voice.GetVoices(string.Empty, string.Empty).Item(3); //其中3为中文，024为英文
                voice.Speak(textspeechText.Text, SpeechVoiceSpeakFlags.SVSFDefault);
                 */

                SpeechVoiceSpeakFlags SpFlags = SpeechVoiceSpeakFlags.SVSFlagsAsync;
                SpVoice Voice = new SpVoice();
                Voice.Speak(this.textspeechText.Text, SpFlags);

            }
            catch (Exception er)
            {
                //MessageBox.Show("An Error Occured!", "SpeechApp", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Console.WriteLine(er.ToString());
            }
            
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            try
            {
                SpeechVoiceSpeakFlags SpFlags = SpeechVoiceSpeakFlags.SVSFDefault;
                SpVoice Voice = new SpVoice();
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "All files (*.*)|*.*|wav files (*.wav)|*.wav";
                sfd.Title = "Save to a wave file";
                sfd.FilterIndex = 2;
                sfd.RestoreDirectory = true;
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    SpeechStreamFileMode SpFileMode = SpeechStreamFileMode.SSFMCreateForWrite;
                    SpFileStream SpFileStream = new SpFileStream();
                    SpFileStream.Open(sfd.FileName, SpFileMode, false);
                    Voice.AudioOutputStream = SpFileStream;
                    Voice.Speak(textspeechText.Text, SpFlags);
                    Voice.WaitUntilDone(1000);
                    SpFileStream.Close();
                }
            }
            catch (Exception er)
            {
                //MessageBox.Show("An Error Occured!", "SpeechApp", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Console.WriteLine(er.ToString());
            }
        }

        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {

            _WAVFile = comboBoxEx1.SelectedText;
            if (_WAVFile == null) return;
            wavRecoContext = new SpeechLib.SpInProcRecoContext();

            ((SpInProcRecoContext)wavRecoContext).Recognition +=
            new _ISpeechRecoContextEvents_RecognitionEventHandler(wavRecoContext_Recognition);

            ((SpInProcRecoContext)wavRecoContext).EndStream +=
                new _ISpeechRecoContextEvents_EndStreamEventHandler(wavRecoContext_EndStream);

            Grammar = wavRecoContext.CreateGrammar(0);

            Grammar.DictationLoad("", SpeechLoadOption.SLOStatic);

            InputWAV = new SpFileStream();

            InputWAV.Open(@_WAVFile, SpeechStreamFileMode.SSFMOpenForRead, false);

            wavRecoContext.Recognizer.AudioInputStream = InputWAV;

            Grammar.DictationSetState(SpeechRuleState.SGDSActive);
        }

        private void labelX3_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();

                dialog.Title = "Select a file";

                dialog.ShowDialog();
                string path = dialog.FileName;
                if (path == null) return;
                FileStream fs = new FileStream(path, FileMode.Open);
                StreamReader reader = new StreamReader(fs);
                string sLine = "";
                sLine = reader.ReadLine();
                while (sLine != null)
                {
                    sLine = reader.ReadLine();
                    if (sLine != null)
                        textspeechText.Text += sLine;
                }
                reader.Close();

            }
            catch (Exception er)
            {
                //MessageBox.Show("An Error Occured!", "SpeechApp", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Console.WriteLine(er.ToString());
            }
        }

        private void labelX2_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();

                dialog.Title = "Select a file";

                dialog.ShowDialog();
                lyricpath = dialog.FileName;
                if (lyricpath == null) 
                    return;

                FileStream fs = new FileStream(lyricpath, FileMode.Open);
                StreamReader reader = new StreamReader(fs);
                string sLine = "";
                sLine = reader.ReadLine();
                while (sLine != null)
                {
                    sLine = reader.ReadLine();
                    if (sLine != null)
                        richTextBoxEx1.Text += sLine;
                }
                reader.Close();

            }
            catch (Exception er)
            {
                //MessageBox.Show("An Error Occured!", "SpeechApp", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Console.WriteLine(er.ToString());
            }
        }
        private void buttonX4_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "All files (*.*)|*.*|lrc files (*.wav)|*.wav";
            save.Title = "Save to a lrc file";
            save.FilterIndex = 2;
            save.RestoreDirectory = true;
            if (save.ShowDialog() == DialogResult.OK)
            {
                FileStream SpFileStream = new FileStream(save.FileName,FileMode.Create);
                StreamWriter write = new StreamWriter(SpFileStream);
                write.WriteLine(richTextBoxEx1.Text);
                SpFileStream.Close();
            }
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
            
        }

        
    }
}
