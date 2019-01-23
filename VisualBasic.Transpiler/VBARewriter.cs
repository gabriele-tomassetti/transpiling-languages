using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using static Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory;

namespace VisualBasic
{
    public class VBARewriter : VisualBasicSyntaxRewriter
    {
        private List<String> StructuresWithInitializer ;

        public VBARewriter(List<String> structuresWithInitializer)
        {
            StructuresWithInitializer = structuresWithInitializer;            
        }

        private StatementSyntax CreateInitializer(string identifier, string type, bool isArray)
        {
            if(isArray)
            {                
                return SyntaxFactory.ParseExecutableStatement($"For Index = 0 To {identifier}.Length - 1{Environment.NewLine}" +
                    $"Call {identifier}(Index).Init{type}(){Environment.NewLine}" +
                    $"Next{Environment.NewLine}");
            }
            else
            {
                return SyntaxFactory.ParseExecutableStatement($"Call {identifier}.Init{type}()");
            }
        }

        public override SyntaxNode VisitModuleBlock(ModuleBlockSyntax node)
        {            
            var initInvocations = new SyntaxList<StatementSyntax>();

            var space = SyntaxFactory.SyntaxTrivia(SyntaxKind.WhitespaceTrivia, " ");

            var newline = SyntaxFactory.SyntaxTrivia(SyntaxKind.EndOfLineTrivia, Environment.NewLine);            

            for (int a = 0; a < node.Members.Count; a++)
            {                
                if (node.Members[a].IsKind(SyntaxKind.FieldDeclaration))
                {                                     
                    foreach (var d in (node.Members[a] as FieldDeclarationSyntax).Declarators)
                    {                        
                        if (StructuresWithInitializer.Contains(
                d?.AsClause?.Type().WithoutTrivia().ToFullString()))
                        {                            
                            foreach (var name in d.Names)
                            {
                                if (name.ArrayBounds != null)
                                {
                                    initInvocations = initInvocations.Add(CreateInitializer(name.Identifier.ToFullString().Trim(), d.AsClause.Type().WithoutTrivia().ToString(), true).WithTrailingTrivia(newline));
                                }
                                else
                                {
                                    initInvocations = initInvocations.Add(CreateInitializer(name.Identifier.ToFullString().Trim(), d.AsClause.Type().WithoutTrivia().ToString(), false).WithTrailingTrivia(newline));
                                }                                
                            }
                        }
                    }
                }
            }

            if (initInvocations.Count > 0)
            {           
                var subStart = SyntaxFactory.SubStatement(SyntaxFactory.Identifier("New()").WithLeadingTrivia(space)).WithLeadingTrivia(newline).WithTrailingTrivia(newline);                
                var subEnd = SyntaxFactory.EndSubStatement(
                    SyntaxFactory.Token(SyntaxKind.EndKeyword, "End ").WithLeadingTrivia(newline), SyntaxFactory.Token(SyntaxKind.SubKeyword, "Sub")).WithTrailingTrivia(newline);

                var moduleConstructor = SyntaxFactory.SubBlock(subStart, initInvocations, subEnd);

                node = node.WithMembers(node.Members.Add(moduleConstructor));
            }

            return base.VisitModuleBlock(node);
        }

        public override SyntaxNode VisitMethodBlock(MethodBlockSyntax node)
        {            
            for (int a = 0; a < node.Statements.Count; a++)
            {
                if(node.Statements[a].IsKind(SyntaxKind.LocalDeclarationStatement))
                {
                    var initInvocations = new SyntaxList<StatementSyntax>();

                    foreach (var d in (node.Statements[a] as LocalDeclarationStatementSyntax).Declarators)
                    {                        
                        if (StructuresWithInitializer.Contains(
                d?.AsClause?.Type().WithoutTrivia().ToString()))
                        {                            
                            foreach (var name in d.Names)
                            {
                                if (name.ArrayBounds != null)
                                {
                                    initInvocations = initInvocations.Add(CreateInitializer(name.Identifier.ToFullString().Trim(), d.AsClause.Type().WithoutTrivia().ToString(), true));
                                }
                                else
                                {
                                    initInvocations = initInvocations.Add(CreateInitializer(name.Identifier.ToFullString().Trim(), d.AsClause.Type().WithoutTrivia().ToString(), false));
                                }
                            }
                        }
                    }
                    
                    foreach (var i in initInvocations)
                    {
                        node = node.WithStatements(node.Statements.Insert(a+1, i.WithTrailingTrivia(SyntaxFactory.SyntaxTrivia(SyntaxKind.EndOfLineTrivia, Environment.NewLine))));
                    }                    
                }
            }

            return base.VisitMethodBlock(node);
        }        

        public override SyntaxNode VisitArgumentList(ArgumentListSyntax node)
        {            
            SeparatedSyntaxList<ArgumentSyntax> arguments = node.Arguments;

            for (int i = 0; i < arguments.Count; i++)
            {
                if(arguments[i].IsKind(SyntaxKind.RangeArgument))
                {                    
                    if (arguments[i].GetFirstToken().IsKind(SyntaxKind.IntegerLiteralToken))
                    {
                        if (arguments[i].GetFirstToken().ValueText == "1")
                        {                            
                            var newStart = SyntaxFactory.IntegerLiteralToken("0 ", LiteralBase.Decimal, TypeCharacter.IntegerLiteral, 0);
                            var newArgument = arguments[i].ReplaceToken(arguments[i].GetFirstToken(), newStart);
                            arguments = arguments.Replace(arguments[i], newArgument);
                        }
                    }                      
                }                                
            }

            return node.WithArguments(arguments);
        }              
    }
}
 
 