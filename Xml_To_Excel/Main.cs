using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Xml_To_Excel
{
    public interface IMainForm
    {
        string SelectXmlFolderPath { get; }
        string SelectExelPath { get; }
        string SelectSaveFolderPath { get; }
        string SaveName { get; }

        event EventHandler SelectXml;
        event EventHandler SelectExel;
        event EventHandler SavePath;
        event EventHandler Save;
    }
    public partial class Main : Form, IMainForm
    {
        public Main()
        {
            InitializeComponent();

            btnSelectXml.Click += btnSelectXml_Click;
            btnSelectExel.Click += btnSelectExel_Click;
            btnSavePath.Click += btnSavePath_Click;
            btnSave.Click += btnSave_Click;
        }
        #region IMainForm
        public string SelectExelPath
        {
            get { return tbSelectXml.Text; }
        }
        public string SelectXmlFolderPath
        {
            get { return tbSelectExel.Text; }
        }
        public string SelectSaveFolderPath
        {
            get { return tbSavePath.Text; }
        }
        public string SaveName
        {
            get { return tbSaveName.Text; }
        }

        public event EventHandler SelectExel;
        public event EventHandler SelectXml;
        public event EventHandler SavePath;
        public event EventHandler Save;
        #endregion

        #region Select exel
        private void btnSelectXml_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                tbSelectXml.Text = dlg.SelectedPath;

                SelectXml?.Invoke(this, EventArgs.Empty);
            }
        }
        private void btnSavePath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                tbSavePath.Text = dlg.SelectedPath;

                SavePath?.Invoke(this, EventArgs.Empty);
            }
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            Save?.Invoke(this, EventArgs.Empty);
        }

        private void btnSelectExel_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel|*.xml;*.xlsx|Все файлы|*";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                tbSelectExel.Text = dlg.FileName;

                SelectExel?.Invoke(this, EventArgs.Empty);
            }
        }
        #endregion
    }
}

