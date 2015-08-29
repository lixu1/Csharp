using System;

using System.Collections.Generic;

using System.ComponentModel;

using System.Data;

using System.Drawing;

using System.Linq;

using System.Text;

using System.Windows.Forms;

using SpeechLib;
using System.IO;

namespace SpeechFromFile
{

    public partial class Form1 : Form
    {

        private SpeechLib.ISpeechRecoContext wavRecoContext = null;
        private SpeechLib.SpFileStream InputWAV = null;
        private SpeechLib.ISpeechRecoGrammar Grammar = null;

        private String _WAVFile = null;

        private string strData = "No recording yet";

        private String _lastRecognized = "";

        public Form1()
        {

            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

   

        //***************************************************************

        //' Event fired when speech recognition engine recognizes audio

        //***************************************************************

        private void wavRecoContext_Recognition(int StreamNumber, object StreamPosition, SpeechRecognitionType RecognitionType, ISpeechRecoResult Result)
        {

            strData = Result.PhraseInfo.GetText(0, -1, true);

            // parseSpeechResult(strData); -- Call a function to parse it if wanted

            _lastRecognized = textBox1.Text;

            textBox1.Text = strData;

        }

        //***************************************************************

        //' End of wav Input Stream reached by speech recognition engine

        //***************************************************************

        private void wavRecoContext_EndStream(int StreamNumber, object StreamPosition, bool f)
        {

            // ' Disable dictation

            Grammar.DictationSetState(SpeechRuleState.SGDSInactive);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();

            dialog.Title = "Select a Speech file";

            dialog.ShowDialog();

            _WAVFile = dialog.FileName;

            if (_WAVFile == null) return;

            //***********************************************

            // Now we have the WAV file, we can set up the

            // inline SPeech Engine to process it

            //***********************************************

            // create the recognition context

            wavRecoContext = new SpeechLib.SpInProcRecoContext();

            //******************************************************************

            // Register our event as a listener on the Recognition event

            // that way, anytime the speech engine thinks it "hears" something

            // that it recognize, we're called to check it out for ourselves

            //********************************************************************

            ((SpInProcRecoContext)wavRecoContext).Recognition +=

            new _ISpeechRecoContextEvents_RecognitionEventHandler(wavRecoContext_Recognition);

            //******************************************************************

            // Register a method on the endStream event so we can do basic clean-up

            // when the Audio file is finished

            //******************************************************************

            ((SpInProcRecoContext)wavRecoContext).EndStream += new _ISpeechRecoContextEvents_EndStreamEventHandler(wavRecoContext_EndStream);

            //*************************************************************************

            // the parameter passed to CreateGrammar is any int. It is only used as if

            // you have more than one grammar active, so you can specify which one is

            // to be used....

            //*************************************************************************

            Grammar = wavRecoContext.CreateGrammar(1);

            // I simply use the default Frammar for dictation

            Grammar.DictationLoad("", SpeechLoadOption.SLOStatic);

            //*************************************************************************

            //Speech engine is now ready to go, so set it to the

            // audio file TO do this, we open the requested file

            // using the SPeechStreamFileMode, and pass that to the

            // speech engine to use as its input source

            //*************************************************************************

            InputWAV = new SpFileStream();

            InputWAV.Open(@_WAVFile, SpeechStreamFileMode.SSFMOpenForRead, false);

            wavRecoContext.Recognizer.AudioInputStream = InputWAV;

            //*************************************************************************

            // the way you "Turn On" the speech engine is by setting the Diction State

            // of its grammar to "Active"

            //*************************************************************************

            Grammar.DictationSetState(SpeechRuleState.SGDSActive);

            //*************************************************************************

            // the result will be handled by wavRecoContext_Recognition()

            //*************************************************************************
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }



    }

}