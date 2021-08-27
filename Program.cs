using System;
/*
 * author : Christi@n
 */
namespace Clock
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                int currentMinute = DateTime.Now.Minute;
                System.Threading.Thread.Sleep(1000);

                if (currentMinute == 0 || currentMinute % 15 == 0)
                    TellMeTime();



            }
        }

        static void startCommandProccess(string command, string  kwargs = "")
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = command;
            startInfo.Arguments = kwargs;
            process.StartInfo = startInfo;
            process.Start();
        }


        static void TellMeTime()
        {
            Random n = new Random();
            string[] sentences = { "Christian. il est", " ", "le temps passe. il est " };

            string correct_words = DateTime.Now.Minute == 0 ? 
                DateTime.Now.Hour.ToString() + " heure " :
                DateTime.Now.Hour.ToString() + " heure " + DateTime.Now.Minute.ToString() + " Minutes";

            string sentence = sentences[n.Next(0, 2)]+correct_words;
            string[] Instructions = { "Dim sapi", "Set sapi = createObject(\"sapi.spvoice\")",
                                            "Set sapi.Voice = sapi.GetVoices.Item(3)",
                                            "sapi.Speak \""+sentence+"\"" };
            //write inside the vbs file
            System.IO.File.WriteAllLines("voice.vbs", Instructions);
            //start new process for cmd
            startCommandProccess("voice.vbs");
            //wait for 1 min
            System.Threading.Thread.Sleep(65000);
            //delete the current vbs file
            System.IO.File.Delete("voice.vbs");
        }
    }

}
