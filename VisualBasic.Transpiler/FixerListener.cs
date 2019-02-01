using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using VisualBasic.Parser;

namespace VisualBasic.Transpiler
{
    public class FixerListener : VBABaseListener
    {
        private CommonTokenStream Stream { get; set; }
        private TokenStreamRewriter Rewriter { get; set; }
        private IToken TokenImport { get; set; } = null;        
        private Stack<string> Types { get; set; } = new Stack<string>();
        private Dictionary<string, StructureInitializer> InitStructures { get; set; } = new Dictionary<string, StructureInitializer>();
        public static List<String> StructuresWithInitializer = new List<string>();

        public string FileName { get; private set; }
        public bool MainFile { get; set; } = false;
       
        //private bool InsideType { get; set; } = false;

        public FixerListener(CommonTokenStream stream, string fileName)
        {
            Stream = stream;
            FileName = fileName.LastIndexOf(".") != -1 ? fileName.Substring(0, fileName.LastIndexOf(".")) : fileName;            
            Rewriter = new TokenStreamRewriter(stream);
        }

        // get the name of the file
        public override void EnterAttributeStmt([NotNull] VBAParser.AttributeStmtContext context)
        {            
            if (context.implicitCallStmt_InStmt().GetText() == "VB_Name")
            {
                FileName = context.literal()[0].GetText().Trim('"');
            }
            
            // remove all attributes
            Rewriter.Replace(context.Start, context.Stop, "");
        }

        // determine whether an option is present
        public override void EnterModuleDeclarationsElement([NotNull] VBAParser.ModuleDeclarationsElementContext context)
        {
            // check whether an option is present
            if (context?.moduleOption() != null && !context.moduleOption().IsEmpty)
            {
                TokenImport = context.moduleOption().Stop;                
            }
        }

        // eliminate header text
        public override void ExitModuleHeader([NotNull] VBAParser.ModuleHeaderContext context)
        {
            // remove any header
            if (context != null)
            {
                Rewriter.Replace(context.Start, context.Stop, "");
            }
        }

        // eliminate config text
        public override void ExitModuleConfig([NotNull] VBAParser.ModuleConfigContext context)
        {
            // remove any config
            if (context != null)
            {
                Rewriter.Replace(context.Start, context.Stop, "");
            }
        }
        
        // wrap the VBA code with a Module
        public override void ExitStartRule(VBAParser.StartRuleContext context)
        {
            // if an option is present it must before everything else
            // therefore we have to check where to put the module start

            var baseText = $"Imports System {Environment.NewLine}{Environment.NewLine}Imports Microsoft.VisualBasic{Environment.NewLine}Imports System.Math{Environment.NewLine}Imports System.Linq{Environment.NewLine}Imports System.Collections.Generic{Environment.NewLine}{Environment.NewLine}Module {FileName}{Environment.NewLine}";

            if (TokenImport != null)
            {
                // add imports and module
                Rewriter.InsertAfter(TokenImport, $"{Environment.NewLine}{baseText}");
            }
            else
            {
                // add imports and module
                Rewriter.InsertBefore(context.Start, baseText);
            }

            Rewriter.InsertAfter(context.Stop.StopIndex, $"{Environment.NewLine}End Module");
        }

        // transform a Type in a Structure
        public override void EnterTypeStmt(VBAParser.TypeStmtContext context)
        {
            var typeName = context.ambiguousIdentifier().GetText();
            Types.Push(typeName);
            InitStructures.Add(typeName, new StructureInitializer(typeName));

            Rewriter.Replace(context.TYPE().Symbol, "Structure");
            Rewriter.Replace(context.END_TYPE().Symbol, "End Structure");

            string visibility = context.visibility().GetText();

            foreach(var st in context.typeStmt_Element())
            {
                Rewriter.InsertBefore(st.Start, $"{visibility} ");
            }
        }
        
        // we add initialization Sub for the current Structure
        public override void ExitTypeStmt([NotNull] VBAParser.TypeStmtContext context)
        {
            var currentType = Types.Pop();

            if (InitStructures.ContainsKey(currentType) && InitStructures[currentType].Text.Length > 0)
            {
                Rewriter.InsertBefore(context.Stop, InitStructures[currentType].Text);

                StructuresWithInitializer.Add(currentType);
            }
            else
            {
                InitStructures.Remove(currentType);
            }
        }

        // you cannot initialize elements inside a Structure
        // since VBA Type(s) are transformed in VB.NET Structure(s)
        // we remove the initialization of array
        public override void ExitTypeStmt_Element([NotNull] VBAParser.TypeStmt_ElementContext context)
        {
            var currentType = Types.Peek();

            if (context.subscripts() != null && !context.subscripts().IsEmpty)
            {
                InitStructures[currentType].Add($"ReDim {context.ambiguousIdentifier().GetText()}({context.subscripts().GetText()})");

                StringBuilder commas = new StringBuilder();
                Enumerable.Range(0, context.subscripts().subscript().Length - 1).ToList().ForEach(x => commas.Append(","));
                Rewriter.Replace(context.subscripts().Start, context.subscripts().Stop, $"{commas.ToString()}");
            }
        }

        // add parentheses to calls that do not have them        
        public override void EnterICS_B_MemberProcedureCall([NotNull] VBAParser.ICS_B_MemberProcedureCallContext context)
        {
            if (!context.argsCall().IsEmpty)
            {
                Rewriter.InsertBefore(context.argsCall().Start, "(");
                Rewriter.InsertAfter(context.argsCall().Stop, ")");
            }
        }

        // eliminate "PtrSafe"
        public override void EnterDeclareStmt([NotNull] VBAParser.DeclareStmtContext context)
        {
            if (context.PTRSAFE() != null)
                Rewriter.Replace(context.PTRSAFE().Symbol, "");            
        }        

        // the Erase statement works differently in VBA and VB.Net
        // in VBA, it deletes the array a re-initialize it
        // in VB.Net, it only deletes it
        public override void EnterEraseStmt([NotNull] VBAParser.EraseStmtContext context)
        {
            Rewriter.Replace(context.Start, context.Stop, $"Array.Clear({context.valueStmt()[0].GetText()}, 0, {context.valueStmt()[0].GetText()}.Length)");
        }        

        // we search for the Main Sub
        public override void EnterSubStmt([NotNull] VBAParser.SubStmtContext context)
        {            
            if (context.ambiguousIdentifier().GetText().Trim() == "Main_Run" ||
                context.ambiguousIdentifier().GetText().Trim() == "Main_Sub" ||
                context.ambiguousIdentifier().GetText().Trim() == "Main")
            {
                MainFile = true;

                Rewriter.Replace(context.ambiguousIdentifier().Start, "Main");
                // Some function of VB.Net are culture-aware,
                // this means, for instance, that when parsing a double from a
                // string it searchs for the proper-culture decimal separator (e.g, ',' or '.'). So, we set a culture that ensure
                // that VB.Net uses a decimal separator '.'                
                Rewriter.InsertBefore(context.block().Start, $"{Environment.NewLine}Dim sw As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew(){Environment.NewLine}");
                Rewriter.InsertBefore(context.block().Start, $"{Environment.NewLine}System.Globalization.CultureInfo.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture{Environment.NewLine}");
                // make the program wait at the end                
                Rewriter.InsertBefore(context.block().Stop, $"{Environment.NewLine}Console.WriteLine(\"Press any key to exit the program\"){Environment.NewLine}Console.ReadKey(){Environment.NewLine}");
                Rewriter.InsertBefore(context.block().Stop, $"{Environment.NewLine}sw.Stop(){Environment.NewLine}Console.WriteLine($\"Time elapsed {{sw.Elapsed}}\"){Environment.NewLine}");
            }
        }

        // returns the changed text
        public string GetText()
        {
            return Rewriter.GetText();
        }
    }    
}
