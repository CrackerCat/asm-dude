// The MIT License (MIT)
//
// Copyright (c) 2021 Henk-Jan Lebbink
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

namespace AsmDude.QuickInfo
{
    using System;
    using System.Collections.Generic;
    using System.Collections.Immutable;
    using System.Diagnostics.Contracts;
    using System.IO;
    using System.Reflection;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Documents;
    using System.Windows.Media;
    using AsmDude.SyntaxHighlighting;
    using AsmDude.Tools;
    using AsmTools;
    using Microsoft.VisualStudio.Core.Imaging;
    using Microsoft.VisualStudio.Imaging;
    using Microsoft.VisualStudio.Language.Intellisense;
    using Microsoft.VisualStudio.Language.StandardClassification;
    using Microsoft.VisualStudio.Text;
    using Microsoft.VisualStudio.Text.Adornments;
    using Microsoft.VisualStudio.Text.Tagging;


    /// <summary>
    /// Provides QuickInfo information to be displayed in a text buffer
    /// </summary>
    internal sealed class AsmQuickInfoSource : IAsyncQuickInfoSource //XYZZY NEW
    //internal sealed class AsmQuickInfoSource : IQuickInfoSource //XYZZY OLD
    {
        private readonly ITextBuffer textBuffer_;
        private readonly ITagAggregator<AsmTokenTag> aggregator_;
        private readonly LabelGraph labelGraph_;
        private readonly AsmSimulator asmSimulator_;
        private readonly AsmDudeTools asmDudeTools_;

        private static readonly ImageId AssemblyWarningImageId = new ImageId(new Guid("{ae27a6b0-e345-4288-96df-5eaf394ee369}"), 200);
        private static readonly ImageId Cube_ = KnownMonikers.AbstractCube.ToImageId(); // http://glyphlist.azurewebsites.net/knownmonikers/
        private static readonly ImageId Branch_ = KnownMonikers.Branch.ToImageId();
        private static readonly ImageId Binary_ = KnownMonikers.Binary.ToImageId();
        private static readonly ImageId Type_ = KnownMonikers.Type.ToImageId();


        public object CSharpEditorResources { get; private set; }

        public AsmQuickInfoSource(
                ITextBuffer textBuffer,
                IBufferTagAggregatorFactoryService aggregatorFactory,
                LabelGraph labelGraph,
                AsmSimulator asmSimulator)
        {
            this.textBuffer_ = textBuffer ?? throw new ArgumentNullException(nameof(textBuffer));
            this.aggregator_ = AsmDudeToolsStatic.GetOrCreate_Aggregator(textBuffer, aggregatorFactory);
            this.labelGraph_ = labelGraph ?? throw new ArgumentNullException(nameof(labelGraph));
            this.asmSimulator_ = asmSimulator ?? throw new ArgumentNullException(nameof(asmSimulator));
            this.asmDudeTools_ = AsmDudeTools.Instance;
        }

        public Task<QuickInfoItem> GetQuickInfoItemAsync(IAsyncQuickInfoSession session, CancellationToken cancellationToken)
        {
            AsmDudeToolsStatic.Output_INFO(string.Format(AsmDudeToolsStatic.CultureUI, "{0}:GetQuickInfoItemAsync", this.ToString()));
            var triggerPoint = session.GetTriggerPoint(this.textBuffer_.CurrentSnapshot);
            if (triggerPoint != null)
            {
                string contentType = this.textBuffer_.ContentType.DisplayName;
                if (contentType.Equals(AsmDudePackage.AsmDudeContentType, StringComparison.Ordinal))
                {
                    QuickInfoItem qii;
                    if (false)
                    {
                        var line = triggerPoint.Value.GetContainingLine();
                        var lineSpan = this.textBuffer_.CurrentSnapshot.CreateTrackingSpan(line.Extent, SpanTrackingMode.EdgeInclusive);
                        var text = triggerPoint.Value.GetContainingLine().GetText(); //get whole line of current cursor pos
                        ContainerElement dataElm = new ContainerElement(
                        ContainerElementStyle.Stacked,
                        new ClassifiedTextElement(
                            new ClassifiedTextRun(PredefinedClassificationTypeNames.Keyword, "BLA 1: " + text.ToString())
                        ));
                        qii = new QuickInfoItem(lineSpan, dataElm);
                    }
                    else
                    {
                        qii = this.HandleNew(session, triggerPoint);
                    }
                    return Task.FromResult(qii); //add custom text from above to Quick Info
                }
                else
                {
                    AsmDudeToolsStatic.Output_WARNING(string.Format(AsmDudeToolsStatic.CultureUI, "{0}:AugmentQuickInfoSession; does not have have AsmDudeContentType: but has type {1}", this.ToString(), contentType));
                }
            }
            return Task.FromResult<QuickInfoItem>(null); //do not add anything to Quick Info
        }

        /// <summary>Determine which pieces of Quickinfo content should be displayed</summary>
        //public void AugmentQuickInfoSession(IQuickInfoSession session, IList<object> quickInfoContent, out ITrackingSpan applicableToSpan) //XYZZY OLD
        public void AugmentQuickInfoSession(IAsyncQuickInfoSession session, IList<object> quickInfoContent, out ITrackingSpan applicableToSpan) //XYZZY NEW
        {
            applicableToSpan = null;
            try
            {
                string contentType = this.textBuffer_.ContentType.DisplayName;
                if (contentType.Equals(AsmDudePackage.AsmDudeContentType, StringComparison.Ordinal))
                {
                    this.Handle(session, quickInfoContent, out applicableToSpan);
                }
                else
                {
                    AsmDudeToolsStatic.Output_WARNING(string.Format(AsmDudeToolsStatic.CultureUI, "{0}:AugmentQuickInfoSession; does not have have AsmDudeContentType: but has type {1}", this.ToString(), contentType));
                }
            }
            catch (Exception e)
            {
                AsmDudeToolsStatic.Output_ERROR(string.Format(AsmDudeToolsStatic.CultureUI, "{0}:AugmentQuickInfoSession; e={1}", this.ToString(), e.ToString()));
            }
        }

        #region Private Methods

        public static string GetPerformanceInfo(Mnemonic mnemonic, AsmDudeTools asmDudeTools)
        {
            Contract.Requires(asmDudeTools != null);
            string result = string.Empty;

            if (Settings.Default.PerformanceInfo_On)
            {
                bool first = true;
                string format = "{0,-14}{1,-24}{2,-7}{3,-9}{4,-20}{5,-9}{6,-11}{7,-10}";

                MicroArch selectedMicroarchitures = AsmDudeToolsStatic.Get_MicroArch_Switched_On();
                foreach (PerformanceItem item in asmDudeTools.Performance_Store.GetPerformance(mnemonic, selectedMicroarchitures))
                {
                    if (first)
                    {
                        first = false;
                        result += string.Format(
                            AsmDudeToolsStatic.CultureUI,
                            format,
                            string.Empty, string.Empty, "µOps", "µOps", "µOps", string.Empty, string.Empty, string.Empty);

                        result += string.Format(
                            AsmDudeToolsStatic.CultureUI,
                            "\n" + format,
                            "Architecture", "Instruction", "Fused", "Unfused", "Port", "Latency", "Throughput", string.Empty);
                    }
                    result += string.Format(
                        AsmDudeToolsStatic.CultureUI,
                        "\n" + format,
                        item.microArch_ + " ",
                        item.instr_ + " " + item.args_ + " ",
                        item.mu_Ops_Fused_ + " ",
                        item.mu_Ops_Merged_ + " ",
                        item.mu_Ops_Port_ + " ",
                        item.latency_ + " ",
                        item.throughput_ + " ",
                        item.remark_);
                }
            }
            return result;
        }


        private QuickInfoItem HandleNew(IAsyncQuickInfoSession session, SnapshotPoint? triggerPoint) //XYZZY NEW
        {
            DateTime time1 = DateTime.Now;

            QuickInfoItem result = null;

            IList<object> quickInfoContent = new List<object>();
            ITrackingSpan applicableToSpan = null;
            ITextSnapshot snapshot = this.textBuffer_.CurrentSnapshot;

            Brush foreground = AsmDudeToolsStatic.GetFontColor();

            (AsmTokenTag tag, SnapshotSpan? keywordSpan) = AsmDudeToolsStatic.GetAsmTokenTag(this.aggregator_, triggerPoint.Value);
            if (keywordSpan.HasValue)
            {
                SnapshotSpan tagSpan = keywordSpan.Value;
                string keyword = tagSpan.GetText();
                string keyword_upcase = keyword.ToUpperInvariant();

                AsmDudeToolsStatic.Output_INFO(string.Format(AsmDudeToolsStatic.CultureUI, "{0}:Handle: keyword=\"{1}\"; type={2}; file=\"{3}\"", this.ToString(), keyword, tag.Type, AsmDudeToolsStatic.GetFilename(session.TextView.TextBuffer)));
                applicableToSpan = snapshot.CreateTrackingSpan(tagSpan, SpanTrackingMode.EdgeInclusive);

                TextBlock description = null;
                switch (tag.Type)
                {
                    case AsmTokenType.Misc:
                        {
                            string descr = this.asmDudeTools_.Get_Description(keyword_upcase);
                            if (descr.Length > 0)
                            {
                                if (keyword.Length > (AsmDudePackage.MaxNumberOfCharsInToolTips / 2))
                                {
                                    descr = "\n" + descr;
                                }
                                string full_Descr = AsmSourceTools.Linewrap(": " + descr, AsmDudePackage.MaxNumberOfCharsInToolTips);

                                result = new QuickInfoItem(
                                    applicableToSpan,
                                    new ContainerElement(
                                        ContainerElementStyle.Stacked,
                                        new ContainerElement(
                                            ContainerElementStyle.Wrapped,
                                            new ImageElement(Type_),
                                            new ClassifiedTextElement(
                                                new ClassifiedTextRun(PredefinedClassificationTypeNames.Text, "Keyword "),
                                                new ClassifiedTextRun(PredefinedClassificationTypeNames.Keyword, keyword),
                                                new ClassifiedTextRun(PredefinedClassificationTypeNames.Text, full_Descr)))));
                            }
                            break;
                        }
                    case AsmTokenType.Directive:
                        {
                            string descr = this.asmDudeTools_.Get_Description(keyword_upcase);
                            if (descr.Length > 0)
                            {
                                if (keyword.Length > (AsmDudePackage.MaxNumberOfCharsInToolTips / 2))
                                {
                                    descr = "\n" + descr;
                                }
                                string full_Descr = AsmSourceTools.Linewrap(": " + descr, AsmDudePackage.MaxNumberOfCharsInToolTips);

                                result = new QuickInfoItem(
                                    applicableToSpan,
                                    new ContainerElement(
                                        ContainerElementStyle.Stacked,
                                        new ContainerElement(
                                            ContainerElementStyle.Wrapped,
                                            new ImageElement(Type_),
                                            new ClassifiedTextElement(
                                                new ClassifiedTextRun(PredefinedClassificationTypeNames.Text, "Directive "),
                                                new ClassifiedTextRun(PredefinedClassificationTypeNames.Keyword, keyword),
                                                new ClassifiedTextRun(PredefinedClassificationTypeNames.Text, full_Descr)))));
                            }
                            break;
                        }
                    case AsmTokenType.Register:
                        {
                            int lineNumber = AsmDudeToolsStatic.Get_LineNumber(tagSpan);
                            if (keyword_upcase.StartsWith("%", StringComparison.Ordinal))
                            {
                                keyword_upcase = keyword_upcase.Substring(1); // remove the preceding % in AT&T syntax
                            }

                            Rn reg = RegisterTools.ParseRn(keyword_upcase, true);
                            if (this.asmDudeTools_.RegisterSwitchedOn(reg))
                            {
                                string regStr = reg.ToString();
                                Arch arch = RegisterTools.GetArch(reg);
                                string archStr = (arch == Arch.ARCH_NONE) ? string.Empty : " [" + ArchTools.ToString(arch) + "] ";
                                string descr = this.asmDudeTools_.Get_Description(regStr);

                                if (regStr.Length > (AsmDudePackage.MaxNumberOfCharsInToolTips / 2))
                                {
                                    descr = "\n" + descr;
                                }
                                string full_Descr = AsmSourceTools.Linewrap(":" + archStr + descr, AsmDudePackage.MaxNumberOfCharsInToolTips);

                                result = new QuickInfoItem(
                                    applicableToSpan,
                                    new ContainerElement(
                                        ContainerElementStyle.Stacked,
                                        new ContainerElement(
                                            ContainerElementStyle.Wrapped,
                                            new ImageElement(Binary_),
                                            new ClassifiedTextElement(
                                                new ClassifiedTextRun(PredefinedClassificationTypeNames.Keyword, regStr),
                                                new ClassifiedTextRun(PredefinedClassificationTypeNames.Text, full_Descr)))));
                            }
                            break;
                        }
                    case AsmTokenType.Mnemonic: // intentional fall through
                    case AsmTokenType.MnemonicOff: // intentional fall through
                    case AsmTokenType.Jump:
                        {
                            (Mnemonic mnemonic, _) = AsmSourceTools.ParseMnemonic_Att(keyword_upcase, true);
                            string mnemonicStr = mnemonic.ToString();
                            string archStr = ArchTools.ToString(this.asmDudeTools_.Mnemonic_Store.GetArch(mnemonic)) + " ";
                            string descr = AsmSourceTools.Linewrap(this.asmDudeTools_.Mnemonic_Store.GetDescription(mnemonic), AsmDudePackage.MaxNumberOfCharsInToolTips);

                            if (mnemonicStr.Length > (AsmDudePackage.MaxNumberOfCharsInToolTips / 2))
                            {
                                descr = "\n" + descr;
                            }
                            string full_Descr = AsmSourceTools.Linewrap(":" + archStr + descr, AsmDudePackage.MaxNumberOfCharsInToolTips);

                            result = new QuickInfoItem(
                                applicableToSpan,
                                new ContainerElement(
                                    ContainerElementStyle.Stacked,
                                    new ContainerElement(
                                        ContainerElementStyle.Wrapped,
                                        new ImageElement(Cube_),
                                        new ClassifiedTextElement(
                                            new ClassifiedTextRun(PredefinedClassificationTypeNames.Keyword, mnemonicStr),
                                            new ClassifiedTextRun(PredefinedClassificationTypeNames.Text, full_Descr))),
                                    new ContainerElement(
                                        ContainerElementStyle.Stacked,
                                        new ClassifiedTextElement(
                                            new ClassifiedTextRun(PredefinedClassificationTypeNames.Text, "Performance\n " + GetPerformanceInfo(mnemonic, this.asmDudeTools_))))));
                            break;
                        }
                    case AsmTokenType.Label:
                        {
                            string label = keyword;
                            string labelPrefix = tag.Misc;
                            string full_Qualified_Label = AsmDudeToolsStatic.Make_Full_Qualified_Label(labelPrefix, label, AsmDudeToolsStatic.Used_Assembler);

                            description.Inlines.Add(Make_Run2(full_Qualified_Label, new SolidColorBrush(AsmDudeToolsStatic.ConvertColor(Settings.Default.SyntaxHighlighting_Label))));

                            string descr = this.Get_Label_Description(full_Qualified_Label);
                            if (descr.Length == 0)
                            {
                                descr = this.Get_Label_Description(label);
                            }
                            if (descr.Length > 0)
                            {
                                if (keyword.Length > (AsmDudePackage.MaxNumberOfCharsInToolTips / 2))
                                {
                                    descr = "\n" + descr;
                                }
                                string full_Descr = AsmSourceTools.Linewrap(": " + descr, AsmDudePackage.MaxNumberOfCharsInToolTips);

                                result = new QuickInfoItem(
                                    applicableToSpan,
                                    new ContainerElement(
                                        ContainerElementStyle.Stacked,
                                        new ContainerElement(
                                            ContainerElementStyle.Wrapped,
                                            new ImageElement(Cube_),
                                            new ClassifiedTextElement(
                                                new ClassifiedTextRun(PredefinedClassificationTypeNames.Text, "Label "),
                                                new ClassifiedTextRun(PredefinedClassificationTypeNames.Keyword, full_Qualified_Label),
                                                new ClassifiedTextRun(PredefinedClassificationTypeNames.Text, full_Descr)))));
                            }
                            break;
                        }
                    case AsmTokenType.LabelDef:
                        {
                            string label = keyword;
                            string extra_Tag_Info = tag.Misc;
                            string full_Qualified_Label;
                            if ((extra_Tag_Info != null) && extra_Tag_Info.Equals(AsmTokenTag.MISC_KEYWORD_PROTO, StringComparison.Ordinal))
                            {
                                full_Qualified_Label = label;
                            }
                            else
                            {
                                full_Qualified_Label = AsmDudeToolsStatic.Make_Full_Qualified_Label(extra_Tag_Info, label, AsmDudeToolsStatic.Used_Assembler);
                            }

                            AsmDudeToolsStatic.Output_INFO("AsmQuickInfoSource:AugmentQuickInfoSession: found label def " + full_Qualified_Label);

                            description = new TextBlock();
                            description.Inlines.Add(Make_Run1("Label ", foreground));
                            description.Inlines.Add(Make_Run2(full_Qualified_Label, new SolidColorBrush(AsmDudeToolsStatic.ConvertColor(Settings.Default.SyntaxHighlighting_Label))));

                            string descr = this.Get_Label_Def_Description(full_Qualified_Label, label);
                            if (descr.Length > 0)
                            {
                                if (keyword.Length > (AsmDudePackage.MaxNumberOfCharsInToolTips / 2))
                                {
                                    descr = "\n" + descr;
                                }
                                string full_Descr = AsmSourceTools.Linewrap(": " + descr, AsmDudePackage.MaxNumberOfCharsInToolTips);

                                result = new QuickInfoItem(
                                  applicableToSpan,
                                  new ContainerElement(
                                      ContainerElementStyle.Stacked,
                                      new ContainerElement(
                                          ContainerElementStyle.Wrapped,
                                          new ImageElement(Cube_),
                                          new ClassifiedTextElement(
                                              new ClassifiedTextRun(PredefinedClassificationTypeNames.Text, "Label "),
                                              new ClassifiedTextRun(PredefinedClassificationTypeNames.Keyword, full_Qualified_Label),
                                              new ClassifiedTextRun(PredefinedClassificationTypeNames.Text, full_Descr)))));
                            }
                            break;
                        }
                    case AsmTokenType.Constant:
                        {
                            (bool valid, ulong value, int nBits) = AsmSourceTools.Evaluate_Constant(keyword);
                            string constantStr = valid
                                ? value + "d = " + value.ToString("X", AsmDudeToolsStatic.CultureUI) + "h = " + AsmSourceTools.ToStringBin(value, nBits) + "b"
                                : keyword;

                            result = new QuickInfoItem(
                                 applicableToSpan,
                                 new ContainerElement(
                                     ContainerElementStyle.Stacked,
                                     new ContainerElement(
                                         ContainerElementStyle.Wrapped,
                                         new ImageElement(Cube_),
                                         new ClassifiedTextElement(
                                             new ClassifiedTextRun(PredefinedClassificationTypeNames.Text, "Constant "),
                                             new ClassifiedTextRun(PredefinedClassificationTypeNames.Keyword, constantStr)))));
                            break;
                        }
                    case AsmTokenType.UserDefined1:
                        {
                            string descr = this.asmDudeTools_.Get_Description(keyword_upcase);
                            if (descr.Length > 0)
                            {
                                if (keyword.Length > (AsmDudePackage.MaxNumberOfCharsInToolTips / 2))
                                {
                                    descr = "\n" + descr;
                                }
                                string full_Descr = AsmSourceTools.Linewrap(": " + descr, AsmDudePackage.MaxNumberOfCharsInToolTips);
                                result = new QuickInfoItem(
                                     applicableToSpan,
                                     new ContainerElement(
                                         ContainerElementStyle.Stacked,
                                         new ContainerElement(
                                             ContainerElementStyle.Wrapped,
                                             new ImageElement(Cube_),
                                             new ClassifiedTextElement(
                                                 new ClassifiedTextRun(PredefinedClassificationTypeNames.Text, "User defined 1: "),
                                                 new ClassifiedTextRun(PredefinedClassificationTypeNames.Keyword, keyword)))));
                            }
                            break;
                        }
                    case AsmTokenType.UserDefined2:
                        {
                            string descr = this.asmDudeTools_.Get_Description(keyword_upcase);
                            if (descr.Length > 0)
                            {
                                if (keyword.Length > (AsmDudePackage.MaxNumberOfCharsInToolTips / 2))
                                {
                                    descr = "\n" + descr;
                                }
                                string full_Descr = AsmSourceTools.Linewrap(": " + descr, AsmDudePackage.MaxNumberOfCharsInToolTips);
                                result = new QuickInfoItem(
                                     applicableToSpan,
                                     new ContainerElement(
                                         ContainerElementStyle.Stacked,
                                         new ContainerElement(
                                             ContainerElementStyle.Wrapped,
                                             new ImageElement(Cube_),
                                             new ClassifiedTextElement(
                                                 new ClassifiedTextRun(PredefinedClassificationTypeNames.Text, "User defined 2: "),
                                                 new ClassifiedTextRun(PredefinedClassificationTypeNames.Keyword, keyword)))));
                            }
                            break;
                        }
                    case AsmTokenType.UserDefined3:
                        {
                            string descr = this.asmDudeTools_.Get_Description(keyword_upcase);
                            if (descr.Length > 0)
                            {
                                if (keyword.Length > (AsmDudePackage.MaxNumberOfCharsInToolTips / 2))
                                {
                                    descr = "\n" + descr;
                                }
                                string full_Descr = AsmSourceTools.Linewrap(": " + descr, AsmDudePackage.MaxNumberOfCharsInToolTips);
                                result = new QuickInfoItem(
                                     applicableToSpan,
                                     new ContainerElement(
                                         ContainerElementStyle.Stacked,
                                         new ContainerElement(
                                             ContainerElementStyle.Wrapped,
                                             new ImageElement(Cube_),
                                             new ClassifiedTextElement(
                                                 new ClassifiedTextRun(PredefinedClassificationTypeNames.Text, "User defined 3: "),
                                                 new ClassifiedTextRun(PredefinedClassificationTypeNames.Keyword, keyword)))));
                            }
                            break;
                        }
                    default:
                        //description = new TextBlock();
                        //description.Inlines.Add(makeRun1("Unused tagType " + asmTokenTag.Tag.type));
                        break;
                }
                if (description != null)
                {
                    description.Focusable = true;
                    description.FontSize = AsmDudeToolsStatic.GetFontSize() + 2;
                    description.FontFamily = AsmDudeToolsStatic.GetFontType();
                    //AsmDudeToolsStatic.Output_INFO(string.Format(AsmDudeToolsStatic.CultureUI, "{0}:AugmentQuickInfoSession; setting description fontSize={1}; fontFamily={2}", this.ToString(), description.FontSize, description.FontFamily));
                    quickInfoContent.Add(description);
                }
            }
            //AsmDudeToolsStatic.Output_INFO("AsmQuickInfoSource:AugmentQuickInfoSession: applicableToSpan=\"" + applicableToSpan + "\"; quickInfoContent,Count=" + quickInfoContent.Count);
            AsmDudeToolsStatic.Print_Speed_Warning(time1, "QuickInfo");

            return result;
        }


        //private void Handle(IQuickInfoSession session, IList<object> quickInfoContent, out ITrackingSpan applicableToSpan) //XYZZY OLD
        private void Handle(IAsyncQuickInfoSession session, IList<object> quickInfoContent, out ITrackingSpan applicableToSpan) //XYZZY NEW
        {
            applicableToSpan = null;
            DateTime time1 = DateTime.Now;

            ITextSnapshot snapshot = this.textBuffer_.CurrentSnapshot;
            SnapshotPoint? triggerPoint = session.GetTriggerPoint(snapshot);
            if (!triggerPoint.HasValue)
            {
                AsmDudeToolsStatic.Output_WARNING(string.Format(AsmDudeToolsStatic.CultureUI, "{0}:Handle: trigger point is null", this.ToString()));
                return;
            }

            Brush foreground = AsmDudeToolsStatic.GetFontColor();

            (AsmTokenTag tag, SnapshotSpan? keywordSpan) = AsmDudeToolsStatic.GetAsmTokenTag(this.aggregator_, triggerPoint.Value);
            if (keywordSpan.HasValue)
            {
                SnapshotSpan tagSpan = keywordSpan.Value;
                string keyword = tagSpan.GetText();
                string keyword_upcase = keyword.ToUpperInvariant();

                AsmDudeToolsStatic.Output_INFO(string.Format(AsmDudeToolsStatic.CultureUI, "{0}:Handle: keyword=\"{1}\"; type={2}; file=\"{3}\"", this.ToString(), keyword, tag.Type, AsmDudeToolsStatic.GetFilename(session.TextView.TextBuffer)));
                applicableToSpan = snapshot.CreateTrackingSpan(tagSpan, SpanTrackingMode.EdgeInclusive);

                TextBlock description = null;
                switch (tag.Type)
                {
                    case AsmTokenType.Misc:
                        {
                            description = new TextBlock();
                            description.Inlines.Add(Make_Run1("Keyword ", foreground));
                            description.Inlines.Add(Make_Run2(keyword, new SolidColorBrush(AsmDudeToolsStatic.ConvertColor(Settings.Default.SyntaxHighlighting_Misc))));

                            string descr = this.asmDudeTools_.Get_Description(keyword_upcase);
                            if (descr.Length > 0)
                            {
                                if (keyword.Length > (AsmDudePackage.MaxNumberOfCharsInToolTips / 2))
                                {
                                    descr = "\n" + descr;
                                }

                                description.Inlines.Add(new Run(AsmSourceTools.Linewrap(": " + descr, AsmDudePackage.MaxNumberOfCharsInToolTips))
                                {
                                    Foreground = foreground,
                                });
                            }
                            break;
                        }
                    case AsmTokenType.Directive:
                        {
                            description = new TextBlock();
                            description.Inlines.Add(Make_Run1("Directive ", foreground));
                            description.Inlines.Add(Make_Run2(keyword, new SolidColorBrush(AsmDudeToolsStatic.ConvertColor(Settings.Default.SyntaxHighlighting_Directive))));

                            string descr = this.asmDudeTools_.Get_Description(keyword_upcase);
                            if (descr.Length > 0)
                            {
                                if (keyword.Length > (AsmDudePackage.MaxNumberOfCharsInToolTips / 2))
                                {
                                    descr = "\n" + descr;
                                }

                                description.Inlines.Add(new Run(AsmSourceTools.Linewrap(": " + descr, AsmDudePackage.MaxNumberOfCharsInToolTips))
                                {
                                    Foreground = foreground,
                                });
                            }
                            break;
                        }
                    case AsmTokenType.Register:
                        {
                            int lineNumber = AsmDudeToolsStatic.Get_LineNumber(tagSpan);
                            if (keyword_upcase.StartsWith("%", StringComparison.Ordinal))
                            {
                                keyword_upcase = keyword_upcase.Substring(1); // remove the preceding % in AT&T syntax
                            }

                            Rn reg = RegisterTools.ParseRn(keyword_upcase, true);
                            if (this.asmDudeTools_.RegisterSwitchedOn(reg))
                            {
                                RegisterTooltipWindow registerTooltipWindow = new RegisterTooltipWindow(foreground);
                                registerTooltipWindow.SetDescription(reg, this.asmDudeTools_);
                                registerTooltipWindow.SetAsmSim(this.asmSimulator_, reg, lineNumber, true);
                                quickInfoContent.Add(registerTooltipWindow);
                            }
                            break;
                        }
                    case AsmTokenType.Mnemonic: // intentional fall through
                    case AsmTokenType.MnemonicOff: // intentional fall through
                    case AsmTokenType.Jump:
                        {
                            (Mnemonic mnemonic, _) = AsmSourceTools.ParseMnemonic_Att(keyword_upcase, true);
                            InstructionTooltipWindow instructionTooltipWindow = new InstructionTooltipWindow(foreground)
                            {
                                Session = session, // set the owner of this windows such that we can manually close this window
                            };
                            instructionTooltipWindow.SetDescription(mnemonic, this.asmDudeTools_);
                            instructionTooltipWindow.SetPerformanceInfo(mnemonic, this.asmDudeTools_);
                            int lineNumber = AsmDudeToolsStatic.Get_LineNumber(tagSpan);
                            instructionTooltipWindow.SetAsmSim(this.asmSimulator_, lineNumber, true);
                            quickInfoContent.Add(instructionTooltipWindow);
                            break;
                        }
                    case AsmTokenType.Label:
                        {
                            string label = keyword;
                            string labelPrefix = tag.Misc;
                            string full_Qualified_Label = AsmDudeToolsStatic.Make_Full_Qualified_Label(labelPrefix, label, AsmDudeToolsStatic.Used_Assembler);

                            description = new TextBlock();
                            description.Inlines.Add(Make_Run1("Label ", foreground));
                            description.Inlines.Add(Make_Run2(full_Qualified_Label, new SolidColorBrush(AsmDudeToolsStatic.ConvertColor(Settings.Default.SyntaxHighlighting_Label))));

                            string descr = this.Get_Label_Description(full_Qualified_Label);
                            if (descr.Length == 0)
                            {
                                descr = this.Get_Label_Description(label);
                            }
                            if (descr.Length > 0)
                            {
                                if (keyword.Length > (AsmDudePackage.MaxNumberOfCharsInToolTips / 2))
                                {
                                    descr = "\n" + descr;
                                }

                                description.Inlines.Add(new Run(AsmSourceTools.Linewrap(": " + descr, AsmDudePackage.MaxNumberOfCharsInToolTips))
                                {
                                    Foreground = foreground,
                                });
                            }
                            break;
                        }
                    case AsmTokenType.LabelDef:
                        {
                            string label = keyword;
                            string extra_Tag_Info = tag.Misc;
                            string full_Qualified_Label;
                            if ((extra_Tag_Info != null) && extra_Tag_Info.Equals(AsmTokenTag.MISC_KEYWORD_PROTO, StringComparison.Ordinal))
                            {
                                full_Qualified_Label = label;
                            }
                            else
                            {
                                full_Qualified_Label = AsmDudeToolsStatic.Make_Full_Qualified_Label(extra_Tag_Info, label, AsmDudeToolsStatic.Used_Assembler);
                            }

                            AsmDudeToolsStatic.Output_INFO("AsmQuickInfoSource:AugmentQuickInfoSession: found label def " + full_Qualified_Label);

                            description = new TextBlock();
                            description.Inlines.Add(Make_Run1("Label ", foreground));
                            description.Inlines.Add(Make_Run2(full_Qualified_Label, new SolidColorBrush(AsmDudeToolsStatic.ConvertColor(Settings.Default.SyntaxHighlighting_Label))));

                            string descr = this.Get_Label_Def_Description(full_Qualified_Label, label);
                            if (descr.Length > 0)
                            {
                                if (keyword.Length > (AsmDudePackage.MaxNumberOfCharsInToolTips / 2))
                                {
                                    descr = "\n" + descr;
                                }

                                description.Inlines.Add(new Run(AsmSourceTools.Linewrap(": " + descr, AsmDudePackage.MaxNumberOfCharsInToolTips))
                                {
                                    Foreground = foreground,
                                });
                            }
                            break;
                        }
                    case AsmTokenType.Constant:
                        {
                            (bool valid, ulong value, int nBits) = AsmSourceTools.Evaluate_Constant(keyword);
                            string constantStr = valid
                                ? value + "d = " + value.ToString("X", AsmDudeToolsStatic.CultureUI) + "h = " + AsmSourceTools.ToStringBin(value, nBits) + "b"
                                : keyword;


                            if (false) // experiment to get text selectable
                            {
                                TextBoxWindow myWindow = new TextBoxWindow();
                                myWindow.MouseRightButtonUp += this.MyWindow_MouseRightButtonUp;
                                myWindow.MyContent.Text = "Constant X: " + constantStr;
                                myWindow.MyContent.Foreground = foreground;
                                myWindow.MyContent.MouseRightButtonUp += this.MyContent_MouseRightButtonUp;
                                quickInfoContent.Add(myWindow);
                            }
                            else
                            {
                                description = new SelectableTextBlock();
                                description.Inlines.Add(Make_Run1("Constant ", foreground));

                                description.Inlines.Add(Make_Run2(constantStr, new SolidColorBrush(AsmDudeToolsStatic.ConvertColor(Settings.Default.SyntaxHighlighting_Constant))));
                            }
                            break;
                        }
                    case AsmTokenType.UserDefined1:
                        {
                            description = new TextBlock();
                            description.Inlines.Add(Make_Run1("User defined 1: ", foreground));
                            description.Inlines.Add(Make_Run2(keyword, new SolidColorBrush(AsmDudeToolsStatic.ConvertColor(Settings.Default.SyntaxHighlighting_Userdefined1))));

                            string descr = this.asmDudeTools_.Get_Description(keyword_upcase);
                            if (descr.Length > 0)
                            {
                                if (keyword.Length > (AsmDudePackage.MaxNumberOfCharsInToolTips / 2))
                                {
                                    descr = "\n" + descr;
                                }

                                description.Inlines.Add(new Run(AsmSourceTools.Linewrap(": " + descr, AsmDudePackage.MaxNumberOfCharsInToolTips))
                                {
                                    Foreground = foreground,
                                });
                            }
                            break;
                        }
                    case AsmTokenType.UserDefined2:
                        {
                            description = new TextBlock();
                            description.Inlines.Add(Make_Run1("User defined 2: ", foreground));
                            description.Inlines.Add(Make_Run2(keyword, new SolidColorBrush(AsmDudeToolsStatic.ConvertColor(Settings.Default.SyntaxHighlighting_Userdefined2))));

                            string descr = this.asmDudeTools_.Get_Description(keyword_upcase);
                            if (descr.Length > 0)
                            {
                                if (keyword.Length > (AsmDudePackage.MaxNumberOfCharsInToolTips / 2))
                                {
                                    descr = "\n" + descr;
                                }

                                description.Inlines.Add(new Run(AsmSourceTools.Linewrap(": " + descr, AsmDudePackage.MaxNumberOfCharsInToolTips))
                                {
                                    Foreground = foreground,
                                });
                            }
                            break;
                        }
                    case AsmTokenType.UserDefined3:
                        {
                            description = new TextBlock();
                            description.Inlines.Add(Make_Run1("User defined 3: ", foreground));
                            description.Inlines.Add(Make_Run2(keyword, new SolidColorBrush(AsmDudeToolsStatic.ConvertColor(Settings.Default.SyntaxHighlighting_Userdefined3))));

                            string descr = this.asmDudeTools_.Get_Description(keyword_upcase);
                            if (descr.Length > 0)
                            {
                                if (keyword.Length > (AsmDudePackage.MaxNumberOfCharsInToolTips / 2))
                                {
                                    descr = "\n" + descr;
                                }

                                description.Inlines.Add(new Run(AsmSourceTools.Linewrap(": " + descr, AsmDudePackage.MaxNumberOfCharsInToolTips))
                                {
                                    Foreground = foreground,
                                });
                            }
                            break;
                        }
                    default:
                        //description = new TextBlock();
                        //description.Inlines.Add(makeRun1("Unused tagType " + asmTokenTag.Tag.type));
                        break;
                }
                if (description != null)
                {
                    description.Focusable = true;
                    description.FontSize = AsmDudeToolsStatic.GetFontSize() + 2;
                    description.FontFamily = AsmDudeToolsStatic.GetFontType();
                    //AsmDudeToolsStatic.Output_INFO(string.Format(AsmDudeToolsStatic.CultureUI, "{0}:AugmentQuickInfoSession; setting description fontSize={1}; fontFamily={2}", this.ToString(), description.FontSize, description.FontFamily));
                    quickInfoContent.Add(description);
                }
            }
            //AsmDudeToolsStatic.Output_INFO("AsmQuickInfoSource:AugmentQuickInfoSession: applicableToSpan=\"" + applicableToSpan + "\"; quickInfoContent,Count=" + quickInfoContent.Count);
            AsmDudeToolsStatic.Print_Speed_Warning(time1, "QuickInfo");
        }

        private void MyWindow_MouseRightButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            //throw new NotImplementedException();
            //e.Handled = true;
        }

        private void MyContent_MouseRightButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // throw new NotImplementedException();
            //e.Handled = true;
        }

        private static Run Make_Run1(string str, Brush foreground)
        {
            return new Run(str)
            {
                Focusable = true,
                FontWeight = FontWeights.Bold,
                Foreground = foreground,
            };
        }

        private static Run Make_Run2(string str, Brush foreground)
        {
            return new Run(str)
            {
                Focusable = true,
                FontWeight = FontWeights.Bold,
                Foreground = foreground,
            };
        }

        private string Get_Label_Description(string label)
        {
            if (this.labelGraph_.Enabled)
            {
                StringBuilder sb = new StringBuilder();
                SortedSet<uint> labelDefs = this.labelGraph_.Get_Label_Def_Linenumbers(label);
                if (labelDefs.Count > 1)
                {
                    sb.AppendLine(string.Empty);
                }
                foreach (uint id in labelDefs)
                {
                    int lineNumber = LabelGraph.Get_Linenumber(id);
                    string filename = Path.GetFileName(this.labelGraph_.Get_Filename(id));
                    string lineContent = (LabelGraph.Is_From_Main_File(id))
                        ? " :" + this.textBuffer_.CurrentSnapshot.GetLineFromLineNumber(lineNumber).GetText()
                        : string.Empty;
                    sb.AppendLine(AsmDudeToolsStatic.Cleanup(string.Format(AsmDudeToolsStatic.CultureUI, "Defined at LINE {0} ({1}){2}", lineNumber + 1, filename, lineContent)));
                }
                string result = sb.ToString();
                return result.TrimEnd(Environment.NewLine.ToCharArray());
            }
            else
            {
                return "Label analysis is disabled";
            }
        }

        private string Get_Label_Def_Description(string full_Qualified_Label, string label)
        {
            if (!this.labelGraph_.Enabled)
            {
                return "Label analysis is disabled";
            }

            SortedSet<uint> usage = this.labelGraph_.Label_Used_At_Info(full_Qualified_Label, label);
            if (usage.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                if (usage.Count > 1)
                {
                    sb.AppendLine(string.Empty); // add a newline if multiple usage occurances exist
                }
                foreach (uint id in usage)
                {
                    int lineNumber = LabelGraph.Get_Linenumber(id);
                    string filename = Path.GetFileName(this.labelGraph_.Get_Filename(id));
                    string lineContent;
                    if (LabelGraph.Is_From_Main_File(id))
                    {
                        lineContent = " :" + this.textBuffer_.CurrentSnapshot.GetLineFromLineNumber(lineNumber).GetText();
                    }
                    else
                    {
                        lineContent = string.Empty;
                    }
                    sb.AppendLine(AsmDudeToolsStatic.Cleanup(string.Format(AsmDudeToolsStatic.CultureUI, "Used at LINE {0} ({1}){2}", lineNumber + 1, filename, lineContent)));
                    //AsmDudeToolsStatic.Output_INFO(string.Format(AsmDudeToolsStatic.CultureUI, "{0}:getLabelDefDescription; sb=\"{1}\"", this.ToString(), sb.ToString()));
                }
                string result = sb.ToString();
                return result.TrimEnd(Environment.NewLine.ToCharArray());
            }
            else
            {
                return "Not used";
            }
        }

        #endregion Private Methods

        #region IDisposable Support

        public void Dispose()
        {
            AsmDudeToolsStatic.Output_INFO(string.Format(AsmDudeToolsStatic.CultureUI, "{0}:Dispose", this.ToString()));
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~AsmQuickInfoSource()
        {
            this.Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                // free managed resources
                this.aggregator_.Dispose();
            }
            // free native resources if there are any.
        }
    }
    #endregion

    internal class TextEditorWrapper
    {
        private static readonly Type TextEditorType = Type.GetType("System.Windows.Documents.TextEditor, PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35");
        private static readonly PropertyInfo IsReadOnlyProp = TextEditorType.GetProperty("IsReadOnly", BindingFlags.Instance | BindingFlags.NonPublic);
        private static readonly PropertyInfo TextViewProp = TextEditorType.GetProperty("TextView", BindingFlags.Instance | BindingFlags.NonPublic);
        private static readonly MethodInfo RegisterMethod = TextEditorType.GetMethod(
            "RegisterCommandHandlers",
            BindingFlags.Static | BindingFlags.NonPublic, null, new[] { typeof(Type), typeof(bool), typeof(bool), typeof(bool) }, null);

        private static readonly Type TextContainerType = Type.GetType("System.Windows.Documents.ITextContainer, PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35");
        private static readonly PropertyInfo TextContainerTextViewProp = TextContainerType.GetProperty("TextView");

        private static readonly PropertyInfo TextContainerProp = typeof(TextBlock).GetProperty("TextContainer", BindingFlags.Instance | BindingFlags.NonPublic);

        public static void RegisterCommandHandlers(Type controlType, bool acceptsRichContent, bool readOnly, bool registerEventListeners)
        {
            RegisterMethod.Invoke(null, new object[] { controlType, acceptsRichContent, readOnly, registerEventListeners });
        }

        public static TextEditorWrapper CreateFor(TextBlock tb)
        {
            object textContainer = TextContainerProp.GetValue(tb);

            TextEditorWrapper editor = new TextEditorWrapper(textContainer, tb, false);
            IsReadOnlyProp.SetValue(editor.editor_, true);
            TextViewProp.SetValue(editor.editor_, TextContainerTextViewProp.GetValue(textContainer));

            return editor;
        }

        private readonly object editor_;

        public TextEditorWrapper(object textContainer, FrameworkElement uiScope, bool isUndoEnabled)
        {
            this.editor_ = Activator.CreateInstance(TextEditorType, BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.CreateInstance,
                null, new[] { textContainer, uiScope, isUndoEnabled }, null);
        }
    }

    public class SelectableTextBlock : TextBlock
    {
        static SelectableTextBlock()
        {
            FocusableProperty.OverrideMetadata(typeof(SelectableTextBlock), new FrameworkPropertyMetadata(true));
            TextEditorWrapper.RegisterCommandHandlers(typeof(SelectableTextBlock), true, true, true);

            // remove the focus rectangle around the control
            FocusVisualStyleProperty.OverrideMetadata(typeof(SelectableTextBlock), new FrameworkPropertyMetadata((object)null));
        }

        private readonly TextEditorWrapper editor_;

        public SelectableTextBlock()
        {
            this.editor_ = TextEditorWrapper.CreateFor(this);
        }
    }
}