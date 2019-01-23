using System;
using System.Collections.Generic;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using System.Reflection;
using System.IO;
using Microsoft.CodeAnalysis.Emit;
using System.Linq;
using Antlr4.Runtime;
using VisualBasic.Parser;
using Antlr4.Runtime.Tree;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using Microsoft.CodeAnalysis.MSBuild;

namespace VisualBasic.Transpiler
{
    class Program
    {
        private static List<(string file, string text)> SourceCode = new List<(string file, string text)>();        
        private static string BasePath {
            get {
                string path = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory + "../..");
                
                return path;
            }
        }
        private static string NameProgram = "vbProgram";

        public static Object Locker = new Object();

        public static void Main(string[] args)
        {
            string dir = $"{BasePath}/SampleCode/";

            Directory.CreateDirectory($"{BasePath}/transpilation_output/");
                        
            // delete any previous transpilation artifact
            if (Directory.Exists($"{BasePath}/transpilation_output/transpiled_files/"))
                Directory.Delete($"{BasePath}/transpilation_output/transpiled_files/", true);

            Console.WriteLine($"=== Starting Transpilation ===");
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("===============================");
            Console.WriteLine("== Starting Parallel Parsing ==");
            Console.WriteLine("===============================");
            Console.WriteLine(Environment.NewLine);

            Stopwatch sw = Stopwatch.StartNew();

            // parallel parsing
            List<Error> errors = new List<Error>();

            Parallel.ForEach(Directory.EnumerateFiles(dir), basFile =>
            {
                ParseFile(basFile, ref errors);
            });

            if (errors.Count > 0)
            {
                Console.WriteLine("The following errors were found: ");
                foreach (var error in errors)
                {
                    Console.WriteLine(error.Msg);
                }
            }
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("=============================");
            Console.WriteLine("== Ending Parallel Parsing ==");
            Console.WriteLine("=============================");
            Console.WriteLine(Environment.NewLine);

            TranspileCode(FixerListener.StructuresWithInitializer);

            CompileCode(NameProgram);

            sw.Stop();

            Console.WriteLine("=== Process of Transpilation completed ===");
            Console.WriteLine($"Time elapsed {sw.Elapsed}");
        }

        private static void ParseFile(string basFile, ref List<Error> errors)
        {            
            Console.WriteLine($"Parsing {Path.GetFileName(basFile)}");

            ICharStream inputStream = CharStreams.fromPath(basFile);
            VBALexer lexer = new VBALexer(inputStream);
            CommonTokenStream tokens = new CommonTokenStream(lexer);
            VBAParser parser = new VBAParser(tokens);

            // we remove the standard error listeners and add our own
            parser.RemoveErrorListeners();
            VBAErrorListener errorListener = new VBAErrorListener(ref errors);
            parser.AddErrorListener(errorListener);
            
            lexer.RemoveErrorListeners();
            lexer.AddErrorListener(errorListener);

            var tree = parser.startRule();

            Console.WriteLine($"Starting transpilation of {Path.GetFileName(basFile)}");

            FixerListener listener = new FixerListener(tokens, Path.GetFileName(basFile));
            ParseTreeWalker.Default.Walk(listener, tree);

            if (listener.MainFile == true)
                NameProgram = listener.FileName;

            // we ensure that addition to the list are parallel-safe
            lock (Locker)
            {
                SourceCode.Add((file: basFile, text: listener.GetText()));
            }
        }

        private static void CompileCode(string projectName)
        {
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("==========================");
            Console.WriteLine("== Starting compilation ==");
            Console.WriteLine("==========================");
            Console.WriteLine(Environment.NewLine);

            List<SyntaxTree> vbTrees = new List<SyntaxTree>();

            var files = Directory.EnumerateFiles($"{BasePath}\\transpilation_output\\transpiled_files");

            foreach(var f in files)
            {
                vbTrees.Add(VisualBasicSyntaxTree.ParseText(File.ReadAllText(f)).WithFilePath(f));
            }
            
            // gathering the assemblies
            HashSet<MetadataReference> references = new HashSet<MetadataReference>{
                // load essential libraries
                MetadataReference.CreateFromFile(Assembly.Load(new AssemblyName("mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")).Location),
                MetadataReference.CreateFromFile(Assembly.Load(new AssemblyName("System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")).Location),
                MetadataReference.CreateFromFile(Assembly.Load(new AssemblyName("Microsoft.VisualBasic, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")).Location),
                MetadataReference.CreateFromFile(Assembly.Load(new AssemblyName("System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")).Location),
                // load the assemblies needed for the runtime
                 MetadataReference.CreateFromFile(Assembly.Load(new AssemblyName("System.Data, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")).Location),
                MetadataReference.CreateFromFile($"{BasePath}\\Dlls\\System.Data.SQLite.dll")
            };                     

            var options = new VisualBasicCompilationOptions(outputKind: OutputKind.ConsoleApplication, optimizationLevel: OptimizationLevel.Release, platform: Platform.X64, optionInfer: true, optionStrict: OptionStrict.Off, optionExplicit: true,
                concurrentBuild: true, checkOverflow: false, deterministic: true, rootNamespace: projectName, parseOptions: VisualBasicParseOptions.Default);

            // compilation
           var compilation = VisualBasicCompilation.Create(projectName,
                                 vbTrees,
                                 references,
                                 options
                                );
            
            Directory.CreateDirectory($"{BasePath}/transpilation_output/compilation/");

            var emit = compilation.Emit($"{BasePath}/transpilation_output/compilation/{projectName}.exe");            
            
            if (!emit.Success)
            {
                Console.WriteLine("Compilation unsuccessful");
                Console.WriteLine("The following errors were found:");

                foreach (var d in emit.Diagnostics)
                {                    
                    if (d.Severity == DiagnosticSeverity.Error)
                    {
                        Console.WriteLine(d.GetMessage());
                    }
                }

                // we also write the errors in a file
                using (StreamWriter errors = new StreamWriter($"{BasePath}/transpilation_output/compilation/errors.txt"))
                {
                    foreach (var d in emit.Diagnostics)
                    {
                        if (d.Severity == DiagnosticSeverity.Error)
                        {
                            errors.WriteLine($"{{{Path.GetFileName(d.Location?.SourceTree?.FilePath)}}} {d.Location.GetLineSpan().StartLinePosition} {d.GetMessage()}");
                        }
                    }
                }
            }
            else
            {
                Directory.CreateDirectory($"{BasePath}/transpilation_output/compilation/x86");
                Directory.CreateDirectory($"{BasePath}/transpilation_output/compilation/x64");

                // we have to copy the Dlls for the runtime
                foreach (var libFile in Directory.EnumerateFiles($"{BasePath}\\Dlls\\"))
                {
                    File.Copy(libFile, $"{BasePath}\\transpilation_output\\compilation\\{Path.GetFileName(libFile)}", true);
                }
                foreach (var libFile in Directory.EnumerateFiles($"{BasePath}\\Dlls\\x86\\"))
                {
                    File.Copy(libFile, $"{BasePath}\\transpilation_output\\compilation\\x86\\{Path.GetFileName(libFile)}", true);
                }
                foreach (var libFile in Directory.EnumerateFiles($"{BasePath}\\Dlls\\x64\\"))
                {
                    File.Copy(libFile, $"{BasePath}\\transpilation_output\\compilation\\x64\\{Path.GetFileName(libFile)}", true);
                }

                // we copy a SQLite Db           
                File.Copy($"{BasePath}\\Data\\data.db", $"{BasePath}\\transpilation_output\\compilation\\{Path.GetFileName($"{BasePath}\\Data\\data.db")}", true);

                Console.WriteLine(Environment.NewLine);
                Console.WriteLine("========================");
                Console.WriteLine("== Ending Compilation ==");
                Console.WriteLine("========================");
                Console.WriteLine(Environment.NewLine);                
            }       
        }

        private static void TranspileCode(List<String> structuresWithInitializer)
        {
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("==================================================");
            Console.WriteLine("== Starting 2-nd pass of parallel transpilation ==");
            Console.WriteLine("==================================================");
            Console.WriteLine(Environment.NewLine);
            List<SyntaxTree> vbTrees = new List<SyntaxTree>();
            VBARewriter rewriter = new VBARewriter(structuresWithInitializer);  
            
            Parallel.ForEach(SourceCode, sc => {
                Console.WriteLine($"Completing transpilation of {Path.GetFileName(sc.file)}");

                vbTrees.Add(rewriter.Visit(VisualBasicSyntaxTree.ParseText(sc.text).GetCompilationUnitRoot()).SyntaxTree.WithFilePath(sc.file));                
            });           

            vbTrees.Add(VisualBasicSyntaxTree.ParseText(File.ReadAllText($"{BasePath}/Libs/Runtime.vb")).WithFilePath($"{BasePath}/Libs/Runtime.vb"));

            // create the necessary directories            
            Directory.CreateDirectory($"{BasePath}/transpilation_output/transpiled_files/");

            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("============================");
            Console.WriteLine("== Transpilation complete ==");
            Console.WriteLine("============================");
            Console.WriteLine(Environment.NewLine);

            foreach (var vt in vbTrees)
            {
                string fileName = Path.GetFileName(vt.FilePath);

                if (fileName.LastIndexOf(".") != -1)
                    fileName = fileName.Substring(0, fileName.LastIndexOf("."));

                fileName = fileName + ".vb";

                Console.WriteLine($"Writing on disk VB.NET version of {Path.GetFileName(vt.FilePath)}");
                File.WriteAllText($"{BasePath}/transpilation_output/transpiled_files/{fileName}", vt.ToString());
            }
        }        
    }
}
