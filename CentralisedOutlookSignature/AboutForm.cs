using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;

namespace CentralisedOutlookSignature
{
    public partial class AboutForm : Form
    {
        public AboutForm()
        {
            InitializeComponent();

            VersionLabel.Text = Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }

        private void LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start((sender as LinkLabel).Tag as string);
        }
    }
}