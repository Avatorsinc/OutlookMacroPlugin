using Microsoft.Office.Tools.Ribbon;
using System;
using System.Linq;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookMacroPlugin
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void MoveToTreatedSDButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var addIn = Globals.ThisAddIn;
                addIn.MoveSelectedToTreatedSD();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void SignButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var addIn = Globals.ThisAddIn;
                addIn.Sign();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void IBMButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var addIn = Globals.ThisAddIn;
                addIn.IBM();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void EGButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var addIn = Globals.ThisAddIn;
                addIn.EG();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void NCRButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var addIn = Globals.ThisAddIn;
                addIn.NCR();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void WincorButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var addIn = Globals.ThisAddIn;
                addIn.Wincor();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void NDButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var addIn = Globals.ThisAddIn;
                addIn.ND();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void HPEButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var addIn = Globals.ThisAddIn;
                addIn.HPE();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void SIMAButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var addIn = Globals.ThisAddIn;
                addIn.SIMA();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void MIMButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var addIn = Globals.ThisAddIn;
                addIn.MIM();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void LexmarkButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var addIn = Globals.ThisAddIn;
                addIn.Lexmark();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void AteaButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var addIn = Globals.ThisAddIn;
                addIn.Atea();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void RicohButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var addIn = Globals.ThisAddIn;
                addIn.Ricoh();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An error occurred: " + ex.Message);
            }
        }
    }
}
