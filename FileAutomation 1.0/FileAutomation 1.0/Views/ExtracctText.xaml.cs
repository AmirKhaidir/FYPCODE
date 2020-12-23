﻿using System;
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
using System.Windows.Shapes;

namespace FileAutomation_1._0.Views
{
    /// <summary>
    /// Interaction logic for ExtracctText.xaml
    /// </summary>
    public partial class ExtracctText : Window
    {
        public ExtracctText(string text)
        {
            InitializeComponent();

            ExtractedTextBox.Document.Blocks.Clear();
            ExtractedTextBox.Document.Blocks.Add(new System.Windows.Documents.Paragraph(new Run(text)));
        }
    }
}