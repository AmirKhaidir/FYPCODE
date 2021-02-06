using System;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Forms;

namespace FileAutomation_1._0.Views
{
    /// <summary>
    /// Interaction logic for ExtracctText.xaml
    /// </summary>
    public partial class ExtracctText : Window
    {
        string sentence;
        public ExtracctText(string text)
        {
            InitializeComponent();
            sentence = text;
            ExtractedTextBox.Document.Blocks.Clear();
            ExtractedTextBox.Document.Blocks.Add(new System.Windows.Documents.Paragraph(new Run(text)));
        }

        private void save_Click(object sender, RoutedEventArgs e)
        {
            //Create an instance for word app  
            Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();

            //Set animation status for word application  
            winword.ShowAnimation = false;

            //Set status for word application is to be visible or not.  
            winword.Visible = false;

            //Create a missing variable for missing value  
            object missing = System.Reflection.Missing.Value;

            //Create a new document  
            Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

            //adding text to document  
            document.Content.SetRange(0, 0);
            document.Content.Text = sentence + Environment.NewLine;

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Word File (.docx ,.doc)|*.docx;*.doc";
            sfd.InitialDirectory = @"C:\Users\Acer\Documents\automation";

            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                object filename = sfd.FileName;
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                System.Windows.MessageBox.Show("Document Saved successfully !");
            }
        }

        private void close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
