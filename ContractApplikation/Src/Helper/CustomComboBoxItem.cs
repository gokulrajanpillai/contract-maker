using System;

namespace ContractApplikation.Src.Helper
{
    public class CustomComboBoxItem
    {
        public string Text  { get; private set; }
        public object Value { get; private set; }

        public CustomComboBoxItem(String text, Object value)
        {
            this.Text   = text;
            this.Value  = value;
        }

        public override string ToString() => Text;
    }
}
