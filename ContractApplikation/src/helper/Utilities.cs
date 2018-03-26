

using System;
using System.Windows.Forms;

namespace ContractApplikation.Src.Helper
{
    static class Utilities
    {
        public static string FirstLetterToUpperCase(this string s)
        {
            if (string.IsNullOrEmpty(s))
                return string.Empty;
            char[] a = s.ToCharArray();
            a[0] = char.ToUpper(a[0]);
            return new string(a);
        }

        public static TextBox GenerateTextBoxWithNameAndValue(string name, string value)
        {
            TextBox newTextBox = new TextBox();
            newTextBox.Name = name;
            newTextBox.Text = value;
            return newTextBox;
        }


        public static decimal RoundByTwoDecimalPlaces(decimal value)
        {
            return decimal.Round(value, 2, MidpointRounding.AwayFromZero);
        }


        public static string AddCurrencySymbol(string text)
        {
            return text + " " + Constants.CURRENCY_SYMBOL;
        }

        public static void ClearControls(Control.ControlCollection controls)
        {
            if (controls != null)
            {
                foreach (Control control in controls)
                {
                    if (control is TextBox)
                    {
                        (control as TextBox).Clear();
                    }
                    else if (control is ComboBox)
                    {
                        (control as ComboBox).SelectedItem = null;
                    }
                    else if (control is RadioButton)
                    {
                        (control as RadioButton).Checked = false;
                    }
                    else if (control is Control)
                    {
                        ClearControls(control.Controls);
                    }
                }
            }
        } 
    }
}
