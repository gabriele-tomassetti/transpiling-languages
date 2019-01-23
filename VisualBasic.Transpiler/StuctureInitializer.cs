using System;
using System.Collections.Generic;
using System.Text;

namespace VisualBasic.Transpiler
{
    public class StructureInitializer
    {        
        private string Name { get; set; }
        private StringBuilder Body { get; set; } = new StringBuilder();
        public string Text {
            get
            {
                if (Body.Length == 0)
                    return String.Empty;
                else
                    return $"{Body.ToString()}End Sub{Environment.NewLine}";
            }
        }

        public StructureInitializer(string name)
        {
            Name = name;
        }

        public void Add(string text)
        {
            if (Body.Length == 0)
            {
                Body.Append($"Sub Init{Name}(){Environment.NewLine}");
            }

            Body.Append($"{text}{Environment.NewLine}");
        }
    }
}
