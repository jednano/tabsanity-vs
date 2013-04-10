using System;
using Microsoft.VisualStudio.Text.Editor;

namespace TabSanity
{
    class TabOptionsListener : IDisposable
    {
        protected IWpfTextView TextView;
        protected readonly IEditorOptions Options;
        protected bool ConvertTabsToSpaces;
        protected int IndentSize;

        public TabOptionsListener(IWpfTextView textView)
        {
            TextView = textView;
            Options = textView.Options;

            Options.OptionChanged += OnTextViewOptionChanged;
            TextView.Closed += TextViewOnClosed;

            _OnConvertTabsToSpacesOptionChanged();
            OnIndentSizeOptionChanged();
        }

        private void OnTextViewOptionChanged(object sender, EditorOptionChangedEventArgs e)
        {
            switch (e.OptionId)
            {
                case DefaultOptions.ConvertTabsToSpacesOptionName:
                    OnConvertTabsToSpacesOptionChanged();
                    break;

                case DefaultOptions.IndentSizeOptionName:
                    OnIndentSizeOptionChanged();
                    break;
            }
        }

        protected virtual void OnConvertTabsToSpacesOptionChanged()
        {
            _OnConvertTabsToSpacesOptionChanged();
        }

        private void _OnConvertTabsToSpacesOptionChanged()
        {
            ConvertTabsToSpaces = Options.GetOptionValue(DefaultOptions.ConvertTabsToSpacesOptionId);
        }

        private void OnIndentSizeOptionChanged()
        {
            IndentSize = Options.GetOptionValue(DefaultOptions.IndentSizeOptionId);
        }

        private void TextViewOnClosed(object sender, EventArgs eventArgs)
        {
            Dispose();
        }

        public virtual void Dispose()
        {
            TextView.Closed -= TextViewOnClosed;
            Options.OptionChanged -= OnTextViewOptionChanged;
        }
    }
}
