using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;

namespace Jarvis
{
    public partial class voice : Form
    {
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();
        Choices active = new Choices();
        Choices inactive = new Choices();
        SpeechSynthesizer sSynth = new SpeechSynthesizer();
        PromptBuilder pBuilder = new PromptBuilder();
        SpeechRecognitionEngine sRecognize = new SpeechRecognitionEngine();
        String[] cmd;
        String[] icmd;
        private NotifyIcon TrayIcon;

        int ext = 0;

        public voice()
        {
            InitializeComponent();
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.Form1_MouseDown);
            timer1.Interval = 13000;
            timer1.Enabled = true;
            TrayIcon = new NotifyIcon();
            TrayIcon.Text = "Jarvis";
            TrayIcon.Icon = new Icon("C:\\Users\\Arjun Sahani\\Downloads\\jarvis_SfU_icon.ico");
            TrayIcon.Visible = false;
            TrayIcon.DoubleClick += new System.EventHandler(this.notifyIcon1_DoubleClick);
            //axWindowsMediaPlayer1.URL = Application.StartupPath.ToString() + "\\intro\\Intro gratuite.MP4";
            axWindowsMediaPlayer1.URL = @"C:\Users\Arjun Sahani\Documents\Visual Studio 2017\Projects\Jarvis - Copy\Jarvis\Resources\Intro gratuite.MP4";

        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
                this.WindowState = FormWindowState.Normal;
            TrayIcon.Visible = false;
            this.ShowInTaskbar = true;
        }

        private void voice_Load(object sender, EventArgs e)
        {
            cmd = new string[]  { "jarvis open google chrome", "jarvis open camera", "jarvis open notepad", "jarvis open calendar",
                                "who are you", "jarvis sleep","what date is today","navigate to facebook","navigate to google","navigate to twitter","navigate to youtube",
                                "navigate to amazon","navigate to flipkart","jarvis navigate to facebook","jarvis navigate to google","jarvis navigate to twitter",
                                "jarvis navigate to youtube","jarvis navigate to amazon","jarvis navigate to flipkart","jarvis tell me your configuration","jarvis on which os you are running",
                                "jarvis show confidence","jarvis what time is it","what is your name",
                        "tell me your configuration","on which os you are running","show confidence","jarvis open file explorer","how are you jarvis"};
            icmd = new string[] {"jarvis hide commands","jarvis show me commands", "jarvis wakeup", "jarvis shutdown", "yeah i'm sure jarvis" ,"hello jarvis","where are you","jarvis i can't see you",
                                "good morning jarvis","good afternoon jarvis","good evening jarvis","jarvis hide","hide youself jarvis","jarvis where are you"};
            inactive.Add(icmd );
            

        }
        private void Form1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void sRecognize_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            int ex=0;
            confi.Text=e.Result.Confidence.ToString();

            if(pictureBox1.Enabled == false)
            {
                switch (e.Result.Text) { 
                    case "jarvis show me commands":
                        listBox1.Items.Clear();
                        foreach (String c in cmd)
                        {
                            listBox1.Items.Add(c);
                        }
                        foreach (String c in icmd)
                        {
                            listBox1.Items.Add(c);
                        }
                        listBox1.Sorted = true;
                        listgrpbox.Visible = true;

                        break;
                    case "jarvis hide commands":
                        listgrpbox.Visible = false;

                        break;
                    case "jarvis shutdown":
                        sSynth.SpeakAsync("are you sure sir?");
                        ext = 1;
                        break;
                    case "hello jarvis":
                    case "good morning jarvis":
                    case "good afternoon jarvis":
                    case "good evening jarvis":
                    case "jarvis wakeup":
                        if (pictureBox1.Enabled == false)
                        {
                            pictureBox1.Enabled = true;
                            label1.Text = "Online";
                            label1.ForeColor = Color.RoyalBlue;
                            Grammar gr;
                            
                            active.Add(cmd);
                            gr = new Grammar(new GrammarBuilder(active));
                            sRecognize.LoadGrammarAsync(gr);
                            start();


                        }
                        else
                            sSynth.SpeakAsync("i'm already Online sir");
                        return;
                        
                    case "yeah i'm sure jarvis":
                        if (ext == 1)
                            Application.Exit();
                        return;
                    case "jarvis hide":
                    case "hide youself jarvis":
                        this.WindowState = FormWindowState.Minimized;
                        TrayIcon.Visible = true;
                        sSynth.SpeakAsync("ok sir");
                        this.ShowInTaskbar = false;
                        return;
                    case "jarvis where are you":
                    case "where are you":
                    case "jarvis i can't see you":
                        if (this.WindowState == FormWindowState.Minimized)
                            this.WindowState = FormWindowState.Normal;
                        TrayIcon.Visible = false;
                        sSynth.SpeakAsync("i'm here sir");
                        this.ShowInTaskbar = true;
                        return;
                }
                
            }
            if (pictureBox1.Enabled == true)
            {
                switch (e.Result.Text)
                {
                    case "jarvis hide commands":
                        listgrpbox.Visible = false;

                        break;
                    case "yeah i'm sure jarvis":
                        if (ext == 1)
                        {
                            stop();
                            Application.Exit();
                        }
                            return;
                    case "jarvis shutdown":
                        sSynth.SpeakAsync("are you sure?");
                        ext = 1;
                        break;
                    case "jarvis wakeup":
                            sSynth.SpeakAsync("i'm already Online sir");
                        break;
                    case "how are you jarvis":
                        sSynth.SpeakAsync("i'm fine sir");
                        break;
                    case "jarvis open google chrome":
                        sSynth.SpeakAsync("opening google Chrome");
                        Process.Start("chrome.exe");
                        Nav();
                        break;
                    case "jarvis open camera":
                        sSynth.SpeakAsync("opening camera");
                        Process process = new Process();
                        process.StartInfo.FileName = "microsoft.windows.camera:";
                        process.Start();
                        active.Add(new string[] {"close camera", "jarvis close camera"} );
                        Grammar gr;
                        gr = new Grammar(new GrammarBuilder(active));
                        sRecognize.LoadGrammarAsync(gr);
                        break;
                    case "jarvis open notepad":
                        sSynth.SpeakAsync("opening notepad");
                        Process.Start("notepad.exe");
                        active.Add(new string[] { "close notepad", "jarvis close notepad" });
                        
                        gr = new Grammar(new GrammarBuilder(active));
                        sRecognize.LoadGrammarAsync(gr);
                        break;
                    case "who are you":
                        sSynth.SpeakAsync("i'm Jaarvis");
                        break;
                    case "jarvis sleep":
                        sSynth.SpeakAsync("okay sir");
                        pictureBox1.Enabled = false;
                        label1.Text = "Offline";
                        label1.ForeColor = Color.Red;

                        break;
                    case "jarvis what time is it":
                        
                        sSynth.SpeakAsync(DateTime.Now.ToString("h-mm tt"));

                        break;
                    case "what is your name":
                        sSynth.SpeakAsync("jarvis sir");

                        break;
                    case "jarvis what date is today":
                    case "what date is today":
                        sSynth.SpeakAsync(DateTime.Now.ToString("MM-dd-yyyy"));
                        break;
                    case "jarvis navigate to facebook":
                    case "navigate to facebook":
                        Process.Start("chrome.exe", "https://www.facebook.com");
                        Nav();
                        break;
                    case "jarvis navigate to google":
                    case "navigate to google":
                        Process.Start("chrome.exe", "https://www.google.co.in");
                        Nav();
                        break;
                    case "jarvis navigate to twitter":
                    case "navigate to twitter":
                        Process.Start("chrome.exe", "https://twitter.com/");
                        Nav();
                        break;
                    case "jarvis navigate to youtube":
                    case "navigate to youtube":
                        Process.Start("chrome.exe", "https://www.youtube.com");
                        Nav();
                        break;
                    case "jarvis navigate to amazon":
                    case "navigate to amazon":
                        Process.Start("chrome.exe", "https://www.amazon.in");
                        Nav();
                        break;
                    case "jarvis navigate to flipkart":
                    case "navigate to flipkart":
                        Process.Start("chrome.exe", "https://www.flipkart.com/");
                        Nav();
                        break;
                    case "jarvis tell me your configuration":
                    case "tell me your configuration":
                        sSynth.SpeakAsync("its intel core i5 7th gen processor with 3.1 ghiga hertz turbo boost technology, 8gb 2400 megahertz ddr4 random access memory, 1 terabyte of hard disk drive which runs at 5200 Revolutions per minute.");
                        break;
                    case "jarvis on which os you are running":
                    case "on which os you are running":
                        sSynth.SpeakAsync("microsoft's windows operating system");
                        break;
                    case "jarvis show confidence":
                    case "show confidence":
                        confi.Visible = true;
                        break;
                    
                    case "jarvis open file explorer":
                        Process.Start("explorer.exe");
                        ex = ex +1;
                        active.Add(new string[] { "close file explorer", "jarvis close file explore" });
                        gr = new Grammar(new GrammarBuilder(active));
                        sRecognize.LoadGrammarAsync(gr);
                        break;
                    case "jarvis hide":
                    case "hide youself jarvis":
                        this.WindowState = FormWindowState.Minimized;
                        TrayIcon.Visible = true;
                        sSynth.SpeakAsync("ok sir");
                        this.ShowInTaskbar = false;
                        break;
                    case "jarvis where are you":
                        if (this.WindowState == FormWindowState.Minimized)
                            this.WindowState = FormWindowState.Normal;
                        TrayIcon.Visible = false;
                        this.ShowInTaskbar = true;
                        sSynth.SpeakAsync("i'm here sir");
                        break;
                   
                            case "jarvis show me commands":
                        listBox1.Items.Clear();
                        foreach (String c in cmd)
                                {
                                    listBox1.Items.Add(c);
                                }
                                foreach (String c in icmd)
                                {
                                    listBox1.Items.Add(c);
                                }
                                listBox1.Sorted = true;
                        listgrpbox.Visible = true;

                                break;
                        }
                Process[] pname = Process.GetProcessesByName("chrome");
                if (pname.Length != 0)
                    switch(e.Result.Text){
                        case "new window":
                            SendKeys.Send("^n");
                            break;
                        case "close tab":
                            SendKeys.Send("^w");
                            break;
                        case "incognito":
                            SendKeys.Send("^+n");
                            break;
                        case "close chrome":
                            for (int i = 0; i < pname.Length; i++)
                                pname[i].Kill();
                            break;
                        case "new tab":
                            SendKeys.Send("^t");
                            break;
                        case "close":
                            for (int i = 0; i < pname.Length; i++)
                                pname[i].Kill();
                            break;
                            break;
                        case "fullscreen mode":
                            SendKeys.Send("{f11}");
                            break;
                    }

                Process[] pname1 = Process.GetProcessesByName("windowscamera");
                if (pname1.Length != 0)
                switch (e.Result.Text)
                    {
                        case "close camera":
                            pname1[0].Kill();
                            break;
                        case "jarvis close camera":
                            pname1[0].Kill();
                            break;
                    }
                Process[] pname2 = Process.GetProcessesByName("notepad");
                if (pname2.Length != 0)
                    switch (e.Result.Text)
                    {
                        case "close notepad":
                            pname2[0].Kill();
                            break;
                        case "jarvis close notepad":
                            pname2[0].Kill();
                            break;
                    }
                if (ex != 0)
                {
                    Process[] pname3 = Process.GetProcessesByName("explorer");
                    if (pname3.Length != 0)
                        switch (e.Result.Text)
                        {
                            case "close file explorer":
                                SendKeys.Send("%{f4}");
                                ex = ex - 1;
                                break;
                            case "jarvis close file explorer":
                                SendKeys.Send("%{f4}");
                                ex = ex - 1;
                                break;
                        }
                }

            }




        }
        public void start()
        {
            if (DateTime.Now.ToString("tt")=="AM" && int.Parse( DateTime.Now.ToString("hh")) >= 4)
            {
                sSynth.SpeakAsync("Good Morning sir.");
            }else if (DateTime.Now.ToString("tt") == "PM" && int.Parse(DateTime.Now.ToString("hh")) <= 6)
            {
                sSynth.SpeakAsync("Good afternoon sir.");

            }else
            {
                sSynth.SpeakAsync("Good evening sir.");
            }

        }
        public void stop()
        {
            if (DateTime.Now.ToString("tt") == "AM" && int.Parse(DateTime.Now.ToString("hh")) >= 4)
            {
                sSynth.Speak("have a nice day sir");
            }
            else if (DateTime.Now.ToString("tt") == "PM" && int.Parse(DateTime.Now.ToString("hh")) <= 6)
            {
                sSynth.Speak("bye sir");

            }
            else
            {
                sSynth.Speak("Good night sir");
            }

        }
        public void Nav()
        {
            Grammar gr;
            active.Add(new string[] { "new window", "close tab", "incognito", "close chrome", "new tab", "close", "fullscreen mode" });
            gr = new Grammar(new GrammarBuilder(active));
            sRecognize.LoadGrammarAsync(gr);
        }
        private void sRecognize_SpeechRecognized1(object sender, SpeechRecognizedEventArgs e)
        {
            switch (e.Result.Text)
            {
                case "open":
                    sSynth.SpeakAsync("opening google Chrome");
                    Process.Start("chrome.exe");
                    break;
                case "open camera":
                    sSynth.SpeakAsync("opening camera");
                    Process.Start("Photos.exe");
                    break;
                case "open notepad":
                    sSynth.SpeakAsync("opening notepad");
                    Process.Start("notepad.exe");
                    break;
                case "Open calender":
                    sSynth.SpeakAsync("opening calender");
                    Process.Start("calender.exe");
                    break;
                case "who are you?":
                    sSynth.SpeakAsync("i'm Jaarvis");
                    break;


            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            TrayIcon.Visible = true;
            this.ShowInTaskbar = false;

        }

        private void confi_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            axWindowsMediaPlayer1.Visible = false;
            Grammar gr;
            gr = new Grammar(new GrammarBuilder(inactive));
            try
            {
                sRecognize.RequestRecognizerUpdate();
                sRecognize.LoadGrammar(gr);
                sRecognize.SpeechRecognized += sRecognize_SpeechRecognized;
                sRecognize.SetInputToDefaultAudioDevice();
                sRecognize.RecognizeAsync(RecognizeMode.Multiple);




            }

            catch
            {
                return;
            }
        }
        
        private void listBox1_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void listBox1_MouseMove(object sender, MouseEventArgs e)
        {

        }

        private void listBox1_MouseUp(object sender, MouseEventArgs e)
        {

        }

        private void listBox1_DrawItem(object sender, DrawItemEventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            sSynth.SpeakAsync(listBox1.SelectedItem.ToString());
        }
        //public void speak(String tospeak)
        // {
        //    sRecognize.RecognizeAsyncStop();
        //    sSynth.SpeakAsync(tospeak);
        //    sRecognize.RecognizeAsyn();
        // }
    }
}
