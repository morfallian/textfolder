using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Forms;
using Microsoft.Speech.Recognition;



namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        string str = "";
        string[] stream_txt1;
        public Form1()
        {
            InitializeComponent();
            ReadText();
        }


        private void OpenToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            
            string filename = openFileDialog1.FileName;
            
            string fileText = System.IO.File.ReadAllText(filename);
            richTextBox1.Text = fileText;
            Text = filename;
            str = fileText;
        }

        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Save_Txt();
        }

        private void Save_Txt()
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            
            string filename = saveFileDialog1.FileName;
            
            System.IO.File.WriteAllText(filename, richTextBox1.Text);

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (richTextBox1.Text == str)
                e.Cancel = false;
            else
                switch (MessageBox.Show(this, "Произошли изменения в файле, закрыть без сохранения?", "Closing", MessageBoxButtons.YesNo))
                {
                    case DialogResult.No:
                        e.Cancel = true;
                        Save_Txt();
                        str = richTextBox1.Text;
                        break;
                    default:
                        break;
                }
        }
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            var currentSelStart = richTextBox1.SelectionStart;
            var currentSelLength = richTextBox1.SelectionLength;

            richTextBox1.SelectAll();
            richTextBox1.SelectionColor = SystemColors.WindowText;

            var matches = Regex.Matches(richTextBox1.Text, @"\bclass\b");
            foreach (var match in matches.Cast<Match>())
            {
                richTextBox1.Select(match.Index, match.Length);
                richTextBox1.SelectionColor = Color.Blue;
            }

            richTextBox1.Select(currentSelStart, currentSelLength);
            richTextBox1.SelectionColor = SystemColors.WindowText;
        }

        string path = @"C:\Commands\Commands.txt";
        private void ReadText()
        {
            string[] arStr = File.ReadAllLines(path);
            stream_txt1 = arStr;

        }

        static RichTextBox box;
        static TextBox tbox;
        
        static void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs a)
        {
            string Text = a.Result.Text;
            int line;

            if ( a.Result.Confidence > 0.6 )
                switch (Text)
                {
                    case "подключи модуль":
                        box.Focus();
                        box.SelectedText += "import math\n";
                        line = box.Text.Length;
                        box.SelectionStart = line;
                        break;
                    case "выведи":
                        box.Focus();
                        box.SelectedText += "print()";
                        box.SelectionStart = box.Text.Length - 1;
                        break;
                    case "число":
                        box.Focus();
                        box.SelectedText += "const";
                        box.SelectionStart = box.Text.Length;
                        break;
                    case "новая строка":
                        box.Focus();
                        box.SelectedText = "\n";
                        line = box.Text.Length;
                        box.SelectionStart = line;
                        break;
                    case "сделай табуляцию":
                        box.Focus();
                        box.SelectedText = "\t";
                        line = box.Text.Length;
                        box.SelectionStart = line;
                        break;
                    case "сделай три пробела":
                        box.Focus();
                        box.SelectedText = "   ";
                        line = box.Text.Length;
                        box.SelectionStart = line;
                        break;
                    case "шаг назад":
                        box.Undo();
                        break;
                    case "поиск":
                        tbox.Focus();
                        break;
                    case "курсор в конец":
                        line = box.Text.Length;
                        box.SelectionStart = line;
                        break;
                }

        }

        public bool FindText(string find) 
        {

            bool returnValue = false;
            int linenumber = 0;
            int indexToText = -1;
            if (find.Length > 0)
            {


                indexToText = richTextBox1.Find(find);
                if (indexToText >= 0)
                {
                    textBox1.Text = find;
                    linenumber = richTextBox1.GetLineFromCharIndex(indexToText);
                    richTextBox1.SelectionStart = indexToText;
                    richTextBox1.SelectionLength = find.Length;
                    richTextBox1.Select(indexToText, find.Length);
                    richTextBox1.Focus();
                    returnValue = true;

                }

                else
                {

                    MessageBox.Show("Не найдено!");
                    returnValue = false;
                }
            }
            return returnValue;
        }
        private void Form1_Shown(object sender, EventArgs e)
        {
            box = richTextBox1; 

            System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("ru-ru");
            SpeechRecognitionEngine sre = new SpeechRecognitionEngine(ci);
            sre.SetInputToDefaultAudioDevice();

            sre.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sre_SpeechRecognized);


            Choices commands = new Choices();
            commands.Add(new string[] { "подключи модуль", "выведи", "число", "новая строка" });
            commands.Add(stream_txt1);


            GrammarBuilder gb = new GrammarBuilder();
            gb.Culture = ci;
            gb.Append(commands);


            Grammar g = new Grammar(gb);
            sre.LoadGrammar(g);

            sre.RecognizeAsync(RecognizeMode.Multiple);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FindText(textBox1.Text);
        }
    }
}

