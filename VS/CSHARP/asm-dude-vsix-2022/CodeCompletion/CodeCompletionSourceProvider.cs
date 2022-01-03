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

namespace AsmDude
{
    using System.Collections.Generic;
    using System.ComponentModel.Composition;
    using System.Diagnostics.Contracts;
    using AsmDude.Tools;
    using AsyncCompletionSample.JsonElementCompletion;
    using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion;
    using Microsoft.VisualStudio.Text;
    using Microsoft.VisualStudio.Text.Editor;
    using Microsoft.VisualStudio.Text.Operations;
    using Microsoft.VisualStudio.Text.Tagging;
    using Microsoft.VisualStudio.Utilities;

    [Export(typeof(IAsyncCompletionSourceProvider))]
    [ContentType(AsmDudePackage.AsmDudeContentType)]
    [Name("asmCompletion")]
    [TextViewRole(PredefinedTextViewRoles.Document)]
    internal sealed class CodeCompletionSourceProvider : IAsyncCompletionSourceProvider
    {
        private IDictionary<ITextView, IAsyncCompletionSource> cache = new Dictionary<ITextView, IAsyncCompletionSource>();

        [Import]
        private readonly IBufferTagAggregatorFactoryService aggregatorFactory_ = null;

        [Import]
        private readonly ITextDocumentFactoryService docFactory_ = null;

        [Import]
        private readonly IContentTypeRegistryService contentService_ = null;

        [Import]
        private ElementCatalog Catalog;

        [Import]
        private ITextStructureNavigatorSelectorService StructureNavigatorSelector;

        public IAsyncCompletionSource GetOrCreate(ITextView textView)
        {
            Contract.Requires(textView != null);
            //CodeCompletionSource sc()
            //{
            //    LabelGraph labelGraph = AsmDudeToolsStatic.GetOrCreate_Label_Graph(textView.TextBuffer, this.aggregatorFactory_, this.docFactory_, this.contentService_);
            //    AsmSimulator asmSimulator = AsmSimulator.GetOrCreate_AsmSimulator(textView.TextBuffer, this.aggregatorFactory_);
            //    return new CodeCompletionSource(textView.TextBuffer, labelGraph, asmSimulator);
            //}
            //return textView.Properties.GetOrCreateSingletonProperty(sc);

            if (this.cache.TryGetValue(textView, out var itemSource))
            {
                return itemSource;
            }
            var source = new CodeCompletionSource(this.Catalog, this.StructureNavigatorSelector); // opportunity to pass in MEF parts
            textView.Closed += (o, e) => this.cache.Remove(textView); // clean up memory as files are closed
            this.cache.Add(textView, source);
            return source;
        }
    }
}
