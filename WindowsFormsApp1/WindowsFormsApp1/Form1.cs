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
using System.IO;
using System.Diagnostics;
using System.Data.SqlClient;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using NPOI.SS.Formula.Functions;
using System.Media;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {

        SpeechRecognitionEngine _spe = new SpeechRecognitionEngine();
        SpeechRecognitionEngine _speexe = new SpeechRecognitionEngine();
        SpeechRecognitionEngine _speup = new SpeechRecognitionEngine();
        
        SpeechRecognitionEngine _spemute = new SpeechRecognitionEngine();
        SpeechRecognitionEngine _speexeopen = new SpeechRecognitionEngine();
        
        SpeechSynthesizer _sps = new SpeechSynthesizer();
        Random _random = new Random();
        int _timeout = 0;
        OpenFileDialog ofd = new OpenFileDialog();
        Choices speexeChoice = new Choices();
        Choices speupChoice = new Choices();

        SoundPlayer music;

       

        
        IList<string> namelist = new List<string>();
        static int upd;

        int mode = 0;

        //private void speexeGrammar(string name)
        //{


        //    speexeChoice.Add(new string[] { name });
        //    Grammar gm = new Grammar(new GrammarBuilder(speexeChoice));
        //    _speexeopen.RequestRecognizerUpdate();
        //    _speexeopen.LoadGrammarAsync(gm);
        //    _speup.RequestRecognizerUpdate();
        //    _speup.LoadGrammarAsync(gm);


        //}

        //private void speupGrammar(string text)
        //{

        //    speupChoice.Add(new string[] { text });

        //    Grammar gm = new Grammar(new GrammarBuilder(speupChoice));
        //    _speup.LoadGrammarAsync(gm);

        //}













        void splitstring(string s)
        {
            string[] command = (File.ReadAllLines(@"Applist.txt"));
            var myList = new List<string>();


            foreach (string c in command)
            {
                if(c!=s)
                {
                    myList.Add(c);
                  
                }

            }
            System.IO.File.WriteAllLines(@"M:\Speech\WindowsFormsApp1\WindowsFormsApp1\WindowsFormsApp1\bin\Debug\Applist.txt", myList);

        }




        void applicationlist()
        {


            string[] command = (File.ReadAllLines(@"Applist.txt"));
            listcommand.Items.Clear();
            listcommand.SelectionMode = SelectionMode.None;
            listcommand.Visible = true;


            foreach (string c in command)
            {
                listcommand.Items.Add(c);

            }
            button1.Text = "Hide Application List";
        }

       void hideapplist()
       {
                listcommand.Items.Clear();
                listcommand.Visible = false;
                button1.Text = "Application List";
       }
        

        void hidelist()
        {
            if (mode == 0)
            {
                listcommand.Items.Clear();
                listcommand.Visible = false;
                button1.Text = "List";
            }

            else if(mode == 1)
            {
                listcommand.Items.Clear();
                listcommand.Visible = false;
            }

           

        }


        void showlist()
        {

            if (mode == 0)
            {
                string[] command = (File.ReadAllLines(@"TextFile1.txt"));
                listcommand.Items.Clear();
                listcommand.SelectionMode = SelectionMode.None;
                listcommand.Visible = true;

                _sps.SpeakAsync("Here are the listed commands");
                foreach (string c in command)
                {
                    listcommand.Items.Add(c);

                }
                button1.Text = "Hide List";

            }

            else if(mode == 1)
            {
                string[] command = (File.ReadAllLines(@"TextFile2.txt"));
                listcommand.Items.Clear();
                listcommand.SelectionMode = SelectionMode.None;
                listcommand.Visible = true;

                _sps.SpeakAsync("Here are the listed commands");
                foreach (string c in command)
                {
                    listcommand.Items.Add(c);

                }
                

            }
          

            
        }


        void manualmode()
        {
            groupBox1.Visible = true;
            button1.Visible = false;
            button2.Visible = false;
            button3.Visible = false;
            update.Visible = false;
            delete.Visible = false;
            label1.Visible = false;
            label2.Visible = false;
            textBox1.Visible = false;
            textBox2.Visible = false;
            textBox1.Text = null;
            textBox2.Text = null;
            
          
            checkBox1.Text = "Maual Mode";
            
          
            if (mode == 0)
            {
                button1.Visible = true;
                button2.Visible = true;
            }

            else if(mode == 1)
            {
                button1.Visible = true;
                button2.Visible = true;
                button3.Visible = true;
                update.Visible = true;
                delete.Visible = true;
               

            }
            else if(mode == 2)
            {
                button1.Visible = true;
                button2.Visible = true;
                button3.Visible = true;
                label1.Visible = true;
                textBox1.Visible = true;

            }

            else if(mode == 3)
            {
                delete.Visible = true;
                textBox1.Visible = true;
                label1.Visible = true;
                button3.Visible = true;
                button1.Visible = true;

            }

            else if(mode == 4)

            {
                update.Visible = true;
                textBox1.Visible = true;
                label1.Visible = true;
                button3.Visible = true;
                textBox2.Visible = true;
                label2.Visible = true;
                button1.Visible = true;


            }







        }

        void manualdisable()
        {
            groupBox1.Visible = false;
            checkBox1.Text = "Maual Mode Disable";
        }





        void Run(string name)
        {


            try
            {
                //Initialization:
                //Initiating SQL Connection:
                SqlConnection con = new SqlConnection();

                //ConnectionString:
                con.ConnectionString = "data source =LAPTOP-0LTLV8V5;database = MyDatabase;integrated security = SSPI";

                //Generating SQL Query
                SqlCommand command = new SqlCommand("select Address from A1 where callName = '" + name + "'", con);

                //Opening the connection:
                con.Open();

                //Execute SQL Query:
                //SqlDataReader DR = command.ExecuteReader();

                //Binding reader to binding source
                // BindingSource source = new BindingSource();
                //source.DataSource = DR;
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataSet ds = new DataSet();

                adapter.SelectCommand = command;
                adapter.Fill(ds);

                //Binding gridview or control datacsource to binding source:
                //string address = source.ToString();
                //string address = source.ToString();
                //Disconnect
                con.Close();
                string address = (string)ds.Tables[0].Rows[0].ItemArray[0];

                _sps.SpeakAsync("Starting your application.");

                Process.Start(address);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        void updatef(string from, string to)
        {

            try { 


            _sps.SpeakAsync("Type the name you want to give");

            ////speupGrammar(textBox1.Text);
            ///

            speupChoice.Add(new string[] { to });

            Grammar gm = new Grammar(new GrammarBuilder(speupChoice));
            _speup.LoadGrammarAsync(gm);

            //Initiating SQL Connection:
            SqlConnection con = new SqlConnection();


            //ConnectionString:
            con.ConnectionString = "data source = LAPTOP-0LTLV8V5;database = MyDatabase;integrated security = SSPI";

            //Generating SQL Query
            string sql = "UPDATE A1 SET callName = " + " ' " + to + " ' " + "where Name= '" + from+"'";
            using (SqlCommand cmd = new SqlCommand(sql, con))
            {
                //Opening the connection:
                con.Open();

                //cmd.Parameters.Add("@param1", SqlDbType.Int).Value = int.Parse(textBox1.Text);
                //cmd.Parameters.Add("@param2", SqlDbType.VarChar, 50).Value = textBox2.Text;

                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();

                //Disconnect
                con.Close();

                    MessageBox.Show("Update Successful");

            }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                

            }
            
        }



        void deletef(string d)
        {

            try
            {
                //Initiating SQL Connection:
                SqlConnection con = new SqlConnection();


                //ConnectionString:
                con.ConnectionString = "data source = LAPTOP-0LTLV8V5;database = MyDatabase;integrated security = SSPI";

                //Generating SQL Query
                string sql = "DELETE FROM A1 where Name= '" + d+"'";
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    //Opening the connection:
                    con.Open();

                    //cmd.Parameters.Add("@param1", SqlDbType.Int).Value = int.Parse(textBox1.Text);
                    //cmd.Parameters.Add("@param2", SqlDbType.VarChar, 50).Value = textBox2.Text;

                    cmd.CommandType = CommandType.Text;
                    cmd.ExecuteNonQuery();

                    //Disconnect
                    con.Close();
                    MessageBox.Show("Delete Successful");



                }

                namelist.Remove(d);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        


        void backfromupdate()
        {
            mode = 1;
            listcommand.Items.Clear();
            listcommand.Visible = false;

            button1.Text = "ADD";

            update.Text = "Update";
            _sps.SpeakAsync("Back to application manue");

            if (checkBox1.Checked)
            {
                manualmode();

            }
            else if (!checkBox1.Checked)
            {
                manualdisable();
            }

            _speup.RecognizeAsyncCancel();

            _speexe.RecognizeAsync(RecognizeMode.Multiple);
        }



        void add()
        {
            _sps.SpeakAsync("Select the application you want to add");


            ofd.Filter = "EXE|*.exe";
            if (ofd.ShowDialog() == DialogResult.OK) ;
            {

               

                try
                {


                    string[] command = (File.ReadAllLines(@"Applist.txt"));
                    command[command.Length - 1] = Path.GetFileNameWithoutExtension(ofd.FileName) + "\n  ";
                    System.IO.File.WriteAllLines(@"M:\Speech\WindowsFormsApp1\WindowsFormsApp1\WindowsFormsApp1\bin\Debug\Applist.txt", command);


                    _speexeopen.LoadGrammarAsync(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"Applist.txt")))));
                    _speup.LoadGrammarAsync(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"Applist.txt")))));


                    ////speexeChoice.Add(new string[] {name});
                    ////Grammar gm = new Grammar(new GrammarBuilder(speexeChoice));
                    //////_speexeopen.RequestRecognizerUpdate();
                    ////_speexeopen.LoadGrammarAsync(gm);
                    //////_speup.RequestRecognizerUpdate();
                    ////_speup.LoadGrammarAsync(gm);
                    ////namelist.Add(name);


                    SqlConnection con = new SqlConnection();

                    //ConnectionString:
                    con.ConnectionString = "data source = LAPTOP-0LTLV8V5;database = MyDatabase;integrated security = SSPI";

                    //Generating SQL Query
                    string sql = "INSERT INTO A1(Name,callName,Address) VALUES(@param1,@param2,@param3)";
                    using (SqlCommand cmd = new SqlCommand(sql, con))
                    {
                        //Opening the connection:
                        con.Open();

                        cmd.Parameters.Add("@param1", SqlDbType.NVarChar, 100).Value = Path.GetFileNameWithoutExtension(ofd.FileName);
                        cmd.Parameters.Add("@param2", SqlDbType.NVarChar, 100).Value = Path.GetFileNameWithoutExtension(ofd.FileName);
                        cmd.Parameters.Add("@param3", SqlDbType.NVarChar, 1000).Value = ofd.FileName;
                        cmd.CommandType = CommandType.Text;
                        cmd.ExecuteNonQuery();

                        //Disconnect
                        con.Close();
                    }

                    MessageBox.Show("Add Successful");


                    string address = ofd.FileName;

                    //speexeGrammar(name);
                   

                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                   
                }
            }

        }




        void movetoapplication()
        {

            _sps.SpeakAsync("Moving to Application Sector");

            button1.Text = "ADD";
            update.Visible = true;
            delete.Visible = true;
            button2.Text = "Open Application";
            _spe.RecognizeAsyncCancel();
            _speexe.RecognizeAsync(RecognizeMode.Multiple);

        }


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {


            groupBox1.Visible = false;
            
            _spemute.SetInputToDefaultAudioDevice();
            _spemute.LoadGrammarAsync(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"TextFile1.txt")))));
            //_spemute.LoadGrammar(new DictationGrammar());
            _spemute.SpeechRecognized += spemute_SRecognized;



            _speexe.SetInputToDefaultAudioDevice();
            _speexe.LoadGrammarAsync(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"TextFile2.txt")))));

            _speexe.SpeechRecognized += speexe_SRecognized;

            _speup.SetInputToDefaultAudioDevice();
            _speup.LoadGrammarAsync(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"TextFile3.txt")))));

            _speup.SpeechRecognized += speup_SRecognized;

            _speexeopen.SetInputToDefaultAudioDevice();
            _speexeopen.LoadGrammarAsync(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"TextFile4.txt")))));

            _speexeopen.SpeechRecognized += speexeopen_SRecognized;



            _spe.SetInputToDefaultAudioDevice();
            _spe.LoadGrammarAsync(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"TextFile1.txt")))));
            _spe.LoadGrammarAsync(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"Applist.txt")))));
            // _spe.LoadGrammarAsync(new DictationGrammar());
            _spe.RecognizeAsync(RecognizeMode.Multiple);
            _spe.SpeechRecognized += Default_SRecognized;
            _spe.SpeechDetected += Recognizer;
            _sps.SpeakAsync("Welcome");


        }

        private void speexeopen_SRecognized(object sender, SpeechRecognizedEventArgs e)
        {

            
            string speech = e.Result.Text;
            if (speech == "Back")
            {
                mode = 1;
                listcommand.Items.Clear();
                listcommand.Visible = false;

                button2.Text = "Open Application";

                button1.Text = "ADD";

                _sps.SpeakAsync("Back to application manue");


                if (checkBox1.Checked)
                {
                    manualmode();

                }
                else if (!checkBox1.Checked)
                {
                    manualdisable();
                }
                _speexeopen.RecognizeAsyncCancel();

                _speexe.RecognizeAsync(RecognizeMode.Multiple);
            }
            else if (speech == "Manual")
            {
             
                checkBox1.Checked = true;
                manualmode();


            }

            else if (speech == "Disable")
            {
                checkBox1.Checked = false;
                manualdisable();
            }

            else if(speech == "Show list")
            {
                applicationlist();
            }

            else if (speech == "Stop Listening")
            {
                _sps.SpeakAsync("Call me when you need");
                _speexeopen.RecognizeAsyncCancel();
                _spemute.RecognizeAsync(RecognizeMode.Multiple);
            }

            else if(speech == "Hide list")
            {
                hideapplist();
            }


            else
            {
                Run(speech);
                
            }


    }








        private void speup_SRecognized(object sender, SpeechRecognizedEventArgs e)
        {

            if (upd == 1)
            {
                textBox2.Visible = true;
                label2.Visible = true;

                string speech = e.Result.Text;

                if (speech == "Back")
                {

                    backfromupdate();
                }


                else if (speech == "Manual")
                {
                    mode = 0;
                    checkBox1.Checked = true;
                    manualmode();


                }

                else if (speech == "Disable")
                {
                    checkBox1.Checked = false;
                    manualdisable();
                }

                else if (speech == "Show list")
                {
                    applicationlist();
                }

                else if (speech == "Hide list")
                {
                    hideapplist();
                }
                else if (speech == "Stop Listening")
                {
                    _sps.SpeakAsync("Call me when you need");
                    _speup.RecognizeAsyncCancel();
                    _spemute.RecognizeAsync(RecognizeMode.Multiple);
                }

                else
                {
                  

                        updatef(speech, textBox2.Text);
                    
                }
            }
            









            

            else if (upd ==0)
            {
                string speech = e.Result.Text;

                if (speech == "Back")
                {
                    mode = 1;
                    listcommand.Items.Clear();
                    listcommand.Visible = false;

                    delete.Text = "Delete";
                    button1.Text = "ADD";

                    _sps.SpeakAsync("Back to application manue");

                    if (checkBox1.Checked)
                    {
                        manualmode();

                    }
                    else if (!checkBox1.Checked)
                    {
                        manualdisable();
                    }

                    _speup.RecognizeAsyncCancel();

                    _speexe.RecognizeAsync(RecognizeMode.Multiple);
                }


                else if (speech == "Manual")
                {
                    mode = 0;
                    checkBox1.Checked = true;
                    manualmode();


                }

                else if (speech == "Disable")
                {
                    checkBox1.Checked = false;
                    manualdisable();
                }

                else if (speech == "Show list")
                {
                    applicationlist();
                }

                else if (speech == "Hide list")
                {
                    hideapplist();
                }

                else if (speech == "Stop Listening")
                {
                    _sps.SpeakAsync("Call me when you need");
                    _speup.RecognizeAsyncCancel();
                    _spemute.RecognizeAsync(RecognizeMode.Multiple);
                }
                else
                {
                    
                        deletef(speech);
                        
                  

                }
            }


        }









        private void speexe_SRecognized(object sender, SpeechRecognizedEventArgs e)
        {

            string speech = e.Result.Text;
            if (speech == "Where am I now?")
            {
                _sps.SpeakAsync("Application sector.");
            }

            else if (speech == "What can you do?")
            {
                _sps.SpeakAsync("Right now, I can Open application, Add application, Delete application, Update a sub name or shortcut for application");
            }



            else if (speech == "Add application")
            {
                add();
            }



            else if(speech == "Update name")
            {
                _speexe.RecognizeAsyncCancel();
                _speup.RecognizeAsync(RecognizeMode.Multiple);
                mode = 4;
                _sps.SpeakAsync("Which application do you want to update?");
                delete.Visible = false;
                button1.Text = "Application List";
                button2.Visible = false;
                listcommand.Items.Clear();
                listcommand.Visible = false;
                textBox1.Visible = true;
                label1.Text = "What applicaion do you want to Update: type here-";
                label1.Visible = true;
                label2.Visible = true;
                textBox2.Visible = true;
                update.Text = "Update Now";


            }
            else if (speech == "Delete application")
            {
                mode = 3;
                _sps.SpeakAsync("Which application do you want to delete?");
                update.Visible = false;
                button1.Text = "Application List";
                button2.Visible = false;
                listcommand.Items.Clear();
                listcommand.Visible = false;
                textBox1.Visible = true;
                label1.Text = "What applicaion do you want to Delete: type here-";
                label1.Visible = true;
                delete.Text = "Delete Now";
                _speexe.RecognizeAsyncCancel();
                _speup.RecognizeAsync(RecognizeMode.Multiple);


            }
            else if (speech == "Show list")
            {
                _sps.SpeakAsync("Here are the lists of commands.");
                showlist();

            }
            else if(speech == "Hide list")
            {
                hidelist();
            }

            else if(speech == "Open Application")
            {
                mode = 2;
                _sps.SpeakAsync("Which application do you want to open?");
                textBox1.Visible = true;
                button1.Text = "Application List";
                button1.Visible = true;
                update.Visible = false;
                delete.Visible = false;
                label1.Text = "What applicaion do you want to run: type here-";
                label1.Visible = true;
                button2.Text = "Run";
                _speexe.RecognizeAsyncCancel();
                _speexeopen.RecognizeAsync(RecognizeMode.Multiple);

         


            }


            else if (speech == "Back")
            {
                mode = 0;
                listcommand.Items.Clear();
                listcommand.Visible = false;

                button1.Text = "List";
                button2.Text = "Application";

                _speexe.RecognizeAsyncCancel();
                _sps.SpeakAsync("I am in the main menue now.");
                _spe.RecognizeAsync(RecognizeMode.Multiple);

                if (checkBox1.Checked)
                {
                    manualmode();

                }
                else if (!checkBox1.Checked)
                {
                    manualdisable();
                }
            }

            else if (speech == "Manual")
            {
               
                checkBox1.Checked = true;
                manualmode();


            }

            else if (speech == "Disable")
            {
                checkBox1.Checked = false;
                manualdisable();
            }

            else if (speech == "Stop Listening")
            {
                _sps.SpeakAsync("Call me when you need");
                _speexe.RecognizeAsyncCancel();
                _spemute.RecognizeAsync(RecognizeMode.Multiple);
            }



        }












        private void Default_SRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            int randomnum;
            string speech = e.Result.Text;
            music = new SoundPlayer("music.wav");


            if (speech == "Hello")
            {
                _sps.SpeakAsync("Hi Nabil");
            }

            else if (speech == "How are you")
            {
                _sps.SpeakAsync("I am working fine.");
            }
            else if (speech == "What time is it")
            {
                _sps.SpeakAsync(DateTime.Now.ToString("h mm tt"));
            }
            else if(speech == "Play Music")
            {
                
                music.Play();
            }

            else if(speech == "Stop Music")
            {
                music.Stop();
                
            }
            else if (speech == "Stop")
            {
                _sps.SpeakAsyncCancelAll();
                randomnum = _random.Next(1,2);
                if (randomnum == 1)
                {
                    _sps.SpeakAsync("Ok");
                }

                else if (randomnum == 2)
                {
                    _sps.SpeakAsync("I'm sorry");
                }
            }

            else if (speech == "Stop Listening")
            {
                _sps.SpeakAsync("Call me when you need");
                _spe.RecognizeAsyncCancel();
                _spemute.RecognizeAsync(RecognizeMode.Multiple);
            }

            else if (speech == "Show list")
            {
                showlist();
            }

            else if (speech == "Hide List")
            {
                hidelist();
            }

            else if (speech == "Application")
            {

                mode = 1;
                movetoapplication();
                button3.Visible = true;


            }

            else if (speech == "Manual")
            {
                mode = 0;
                checkBox1.Checked = true;
                manualmode();


            }

            else if(speech == "Disable")
            {
                checkBox1.Checked = false;
                manualdisable();
            }

            else if(speech == "Check mail")
            {
                _sps.SpeakAsync("Here you go.");
                System.Diagnostics.Process.Start("https://www.google.com/intl/bn/gmail/about/");
            }
            else if(speech == "Oprn youtube")
            {
                _sps.SpeakAsync("Opening youtube");
                System.Diagnostics.Process.Start("https://www.youtube.com/");
            }

            else if(speech == "Open facebook")
            {
                _sps.SpeakAsync("What a lame choice. Here you go.");
                System.Diagnostics.Process.Start("https://www.facebook.com/");
            }


        }

        



        private void Recognizer(object sender, SpeechDetectedEventArgs e)
        {
            _timeout = 0;

        }



        private void spemute_SRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string speech = e.Result.Text;
            if (speech == "Wake up")
            {
                _spemute.RecognizeAsyncCancel();
                _sps.SpeakAsync("I am here");
                _spe.RecognizeAsync(RecognizeMode.Multiple);
            }
        }



        private void count_Tick(object sender, EventArgs e)
        {
            if (_timeout == 10)
            {
                _spe.RecognizeAsyncCancel();

            }
            else if (_timeout == 11)
            {
                count.Stop();
                _spemute.RecognizeAsync(RecognizeMode.Multiple);
                _timeout = 0;

            }

        }

        private void volume_Scroll(object sender, EventArgs e)
        {
            _sps.Volume = volume.Value;
        }























        private void button1_Click(object sender, EventArgs e)
        {
           
           

            if (button1.Text == "List")
            {
                showlist();

               
            }
            else if(button1.Text == "Hide List")
            {
                hidelist();
                
            }

            else if(button1.Text == "Application List")
            {

                applicationlist();
            }

            else if(button1.Text == "ADD")
            {
                add();
            }

            else if(button1.Text == "Hide Application List")
            {
                hideapplist();
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {

            if (delete.Text == "Delete")
            {
                mode = 3;
                _sps.SpeakAsync("Which application do you want to delete?");
                update.Visible = false;
                button1.Text = "Application List";
                button2.Visible = false;
                listcommand.Items.Clear();
                listcommand.Visible = false;
                textBox1.Visible = true;
                label1.Text = "What applicaion do you want to Delete: type here-";
                label1.Visible = true;
                delete.Text = "Delete Now";
                _speexe.RecognizeAsyncCancel();
                _speup.RecognizeAsync(RecognizeMode.Multiple);

            }

            else if (delete.Text == "Delete Now")
            {
                mode = 3;

                deletef(textBox1.Text);
            }

        }




        private void button2_Click(object sender, EventArgs e)
        {
            listcommand.Items.Clear();
            listcommand.Visible = false;

            if (button2.Text == "Application")
            {
                mode = 1;
                movetoapplication();
                button3.Visible = true;
                
                
            }

            else if(button2.Text == "Open Application")
            {

                mode = 2;
                _sps.SpeakAsync("Which application do you want to open?");
                textBox1.Visible = true;
                button1.Text = "Application List";
                button1.Visible = true;
                update.Visible = false;
                delete.Visible = false;
                label1.Text = "What applicaion do you want to run: type here-";
                label1.Visible = true;
                button2.Text = "Run";
                _speexe.RecognizeAsyncCancel();
                _speexeopen.RecognizeAsync(RecognizeMode.Multiple);

               

            }

            else if(button2.Text == "Run")
            {
                mode = 2;
                try
                {
                    if (textBox1.Text != null)
                    {
                        Run(textBox1.Text);
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void update_Click(object sender, EventArgs e)
        {


            if (update.Text == "Update")
            {
                _speexe.RecognizeAsyncCancel();
                _speup.RecognizeAsync(RecognizeMode.Multiple);
                mode = 4;
                _sps.SpeakAsync("Which application do you want to update?");
                delete.Visible = false;
                button1.Text = "Application List";
                button2.Visible = false;
                listcommand.Items.Clear();
                listcommand.Visible = false;
                textBox1.Visible = true;
                label1.Text = "What applicaion do you want to Update: type here-";
                label1.Visible = true;
                label2.Visible = true;
                textBox2.Visible = true;
                update.Text = "Update Now";
            }
            else if (update.Text == "Update Now")
            {
                mode = 4;
                updatef(textBox1.Text, textBox2.Text);
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

            if (update.Visible == true && delete.Visible == true)
            {
                mode = 0;
                listcommand.Items.Clear();
                listcommand.Visible = false;
                
                button1.Text = "List";
                button2.Text = "Application";
                
                _speexe.RecognizeAsyncCancel();
                _sps.SpeakAsync("I am in the main menue now.");
                _spe.RecognizeAsync(RecognizeMode.Multiple);

                if(checkBox1.Checked)
                {
                    manualmode();

                }
                else if(!checkBox1.Checked)
                {
                    manualdisable();
                }

            }


            else if (update.Visible == true && delete.Visible == false)
            {
                ////mode = 1;
                ////listcommand.Items.Clear();
                ////listcommand.Visible = false;

                ////button1.Text = "ADD";

                ////update.Text = "Update";
                ////_sps.SpeakAsync("Back to application manue");

                ////if (checkBox1.Checked)
                ////{
                ////    manualmode();

                ////}
                ////else if (!checkBox1.Checked)
                ////{
                ////    manualdisable();
                ////}

                ////_speup.RecognizeAsyncCancel();

                ////_speexe.RecognizeAsync(RecognizeMode.Multiple);
                ///
                backfromupdate();

            }

            else if(update.Visible == false && delete.Visible == true)
            {

                mode = 1;
                listcommand.Items.Clear();
                listcommand.Visible = false;
               
                delete.Text = "Delete";
                button1.Text = "ADD";
               
                _sps.SpeakAsync("Back to application manue");

                if (checkBox1.Checked)
                {
                    manualmode();

                }
                else if (!checkBox1.Checked)
                {
                    manualdisable();
                }

                _speup.RecognizeAsyncCancel();

                _speexe.RecognizeAsync(RecognizeMode.Multiple);
            }

            else if(button2.Text == "Run")
            {
                mode = 1;
                listcommand.Items.Clear();
                listcommand.Visible = false;
               
                button2.Text = "Open Application";
               
                button1.Text = "ADD";
               
                _sps.SpeakAsync("Back to application manue");


                if (checkBox1.Checked)
                {
                    manualmode();

                }
                else if (!checkBox1.Checked)
                {
                    manualdisable();
                }
                _speexeopen.RecognizeAsyncCancel();
                
                _speexe.RecognizeAsync(RecognizeMode.Multiple);
            }

          

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox1.Checked)
            {
                manualmode();
            }
            else if(!checkBox1.Checked)
            {
                manualdisable();
            }
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void listcommand_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
