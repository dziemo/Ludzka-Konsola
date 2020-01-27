using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Globalization;
using WindowsInput;
using System.Text.RegularExpressions;

namespace LudzkaKonsolaApp
{
    /// <summary>
    /// Logika interakcji dla klasy MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        static InputSimulator inputSim = new InputSimulator();
        static CultureInfo cultureInfo = new CultureInfo("en-US");
        static SpeechSynthesizer synth = new SpeechSynthesizer();
        static List<Regex> regices = new List<Regex>();
        SpeechRecognitionEngine recognizer = new SpeechRecognitionEngine(cultureInfo);
        static Dictionary<Regex, Action> actions = new Dictionary<Regex, Action>();
        public MainWindow()
        {
            InitializeComponent();
            InitializeRegices();

            Choices commands = new Choices();
            commands.Add(new string[] { "find", "search", "look", "resume", "pause", "stop", "mute", "next track", "play", "pause", "volume up",
                "volume down", "silence", "next track", "previous track", "google", "quit", "game",  "next song", "song", "track", "previous track",
                "previous song", "previous", "exit", "shutdown", "start game", "mail", "email", "e-mail", "home", "homepage", "time", "menu", "youtube" });
            GrammarBuilder gb = new GrammarBuilder(commands);
            gb.Culture = cultureInfo;
            Grammar grammar = new Grammar(gb);
            recognizer.LoadGrammar(grammar);

            recognizer.LoadGrammar(new DictationGrammar());
            recognizer.MaxAlternates = 3;
            recognizer.BabbleTimeout = TimeSpan.FromSeconds(2.0f);
            recognizer.SpeechRecognized +=
              new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized);
            recognizer.SpeechRecognitionRejected +=
                new EventHandler<SpeechRecognitionRejectedEventArgs>(recognizer_SpeechRecognitionRejected);

            recognizer.SetInputToDefaultAudioDevice();

            synth.SetOutputToDefaultAudioDevice();

            synth.SelectVoice(synth.GetInstalledVoices().ToList().Find(x => x.VoiceInfo.Name.Contains("Zira")).VoiceInfo.Name);

            ContentRendered += MainWindow_ContentRendered;
        }
        
        private void InitializeRegices()
        {
            actions.Add(new Regex(@"^(.*?)(\bfind |search for |look for |search |look |google \b)([^$]*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), Action.Find);
            actions.Add(new Regex(@"^(.*?(\bplay|resume\b)[^$]*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), Action.PlayPause);
            actions.Add(new Regex(@"^(.*?(\bpause|stop\b)[^$]*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), Action.PlayPause);
            actions.Add(new Regex(@"^(.*?(\bvolume up\b)[^$]*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), Action.VolUp);
            actions.Add(new Regex(@"^(.*?(\bvolume down\b)[^$]*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), Action.VolDown);
            actions.Add(new Regex(@"^(.*?(\bsilence|mute\b)[^$]*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), Action.Mute);
            actions.Add(new Regex(@"^(.*?(\bnext track|next song|next \b)[^$]*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), Action.NextTrack);
            actions.Add(new Regex(@"^(.*?(\bprevious track|previous song|previous \b)[^$]*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), Action.PrevTrack);
            actions.Add(new Regex(@"^(.*?(\bquit|exit|shutdown\b)[^$]*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), Action.Quit);
            actions.Add(new Regex(@"^(.*?(\bgame|start game\b)[^$]*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), Action.Game);
            //actions.Add(new Regex(@"^(.*?(\byoutube\b)[^$]*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), Action.YouTube);
            actions.Add(new Regex(@"^(.*?(\bmail|email|e-mail\b)[^$]*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), Action.Mail);
            actions.Add(new Regex(@"^(.*?(\bhome|homepage\b)[^$]*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), Action.Home);
            actions.Add(new Regex(@"^(.*?(\bmenu\b)[^$]*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), Action.Menu);
            actions.Add(new Regex(@"^(.*?(\btime\b)[^$]*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), Action.Time);
        }

        private void MainWindow_ContentRendered(object sender, EventArgs e)
        {
            synth.SpeakAsync("Welcome back to the human console");
        }

        private void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            Console.WriteLine("Recognized text: " + e.Result.Text);

            ProcessSpeech(e.Result.Text);

            EnableButton();
        }

        private void recognizer_SpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            synth.SpeakAsync("Sorry I don't understand that, please repeat");
            EnableButton();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            recognizer.RecognizeAsync(RecognizeMode.Multiple);
            DisableButton();
        }

        private void DisableButton ()
        {
            recognizeBtn.IsEnabled = false;
            recognizeBtn.Background = Brushes.Red;
        }

        private void EnableButton ()
        {
            recognizer.RecognizeAsyncStop();
            recognizeBtn.IsEnabled = true;
            recognizeBtn.Background = Brushes.LightGreen;
        }

        private void Search(string query)
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = "/C start www.google.pl/search?q="+query;
            process.StartInfo = startInfo;
            process.Start();
        }

        private void RunGame()
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = "/C \"D:\\Program Files (x86)\\Steam\\Steam.exe\" steam://rungameid/429330 ";
            process.StartInfo = startInfo;
            process.Start();
        }

        private void YouTubeSearch (string query)
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = "/C start www.youtube.com/results?search_query="+query;
            process.StartInfo = startInfo;
            process.Start();
        }

        private void ProcessSpeech(string speech)
        {
            foreach (var action in actions)
            {
                MatchCollection matches = action.Key.Matches(speech);
                if (matches.Count > 0)
                {
                    TakeAction(action.Value, matches);
                    return;
                }
            }
        }

        private void TakeAction (Action action, MatchCollection matches)
        {
            switch (action)
            {
                case Action.PlayPause:
                    inputSim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.MEDIA_PLAY_PAUSE);
                    break;
                case Action.VolUp:
                    inputSim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.VOLUME_UP);
                    break;
                case Action.VolDown:
                    inputSim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.VOLUME_DOWN);
                    break;
                case Action.Mute:
                    inputSim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.VOLUME_MUTE);
                    break;
                case Action.NextTrack:
                    inputSim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.MEDIA_NEXT_TRACK);
                    break;
                case Action.PrevTrack:
                    inputSim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.MEDIA_PREV_TRACK);
                    break;
                case Action.Find:
                    string query = "";
                    foreach (Match match in matches)
                    {
                        foreach (Group g in match.Groups)
                        {
                            Console.WriteLine("GROUP MATCH " + g.ToString());
                        }
                        query += match.Groups[3];
                    }
                    if (query.Length > 0)
                    {
                        query = query.TrimStart(' ');
                        query = query.Replace(' ', '+');
                        Console.WriteLine(query);
                        Search(query);
                    }
                    break;
                case Action.Quit:
                    synth.Speak("Goodbye");
                    Application.Current.Shutdown();
                    break;
                case Action.Game:
                    RunGame();
                    break;
                //case Action.YouTube:
                //    string youtubeQuery = "";
                //    foreach (Match match in matches)
                //    {
                //        foreach (Group g in match.Groups)
                //        {
                //            Console.WriteLine("GROUP MATCH " + g.ToString());
                //        }
                //        youtubeQuery += match.Groups[3];
                //    }
                //    if (youtubeQuery.Length > 0)
                //    {
                //        youtubeQuery = youtubeQuery.TrimStart(' ');
                //        youtubeQuery = youtubeQuery.Replace(' ', '+');
                //        Console.WriteLine(youtubeQuery);
                //        YouTubeSearch(youtubeQuery);
                //    }
                //    break;
                case Action.Time:
                    synth.SpeakAsync("Time is " + DateTime.Now.ToShortTimeString());
                    break;
                case Action.Mail:
                    inputSim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.LAUNCH_MAIL);
                    break;
                case Action.Home:
                    inputSim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.BROWSER_HOME);
                    break;
                case Action.Menu:
                    inputSim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.MENU);
                    break;
                default:
                    synth.SpeakAsync("I can't do that");
                    break;
            }
        }
        
        public enum Action { PlayPause, VolUp, VolDown, Mute, NextTrack, PrevTrack, Find, Quit, Game, Time, Mail, Home, Menu };
    }
}
