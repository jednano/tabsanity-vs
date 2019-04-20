using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text.Editor;

namespace TabSanity
{
	internal sealed class DisplayWindowHelper
	{
		private ICompletionBroker _completionBroker;
		private ISignatureHelpBroker _signatureHelpBroker;
		private ILightBulbBroker _smartTagBroker;
		private IAsyncQuickInfoBroker _quickInfoBroker;

		private DisplayWindowHelper(
			ITextView view,
			ICompletionBroker completionBroker,
			ISignatureHelpBroker signatureHelpBroker,
			ILightBulbBroker smartTagBroker,
			IAsyncQuickInfoBroker quickInfoBroker)
			: this(completionBroker, signatureHelpBroker, smartTagBroker, quickInfoBroker)
		{
			TextView = view;
		}

		internal DisplayWindowHelper(
			ICompletionBroker completionBroker,
			ISignatureHelpBroker signatureHelpBroker,
			ILightBulbBroker smartTagBroker,
			IAsyncQuickInfoBroker quickInfoBroker)
		{
			_completionBroker = completionBroker;
			_signatureHelpBroker = signatureHelpBroker;
			_smartTagBroker = smartTagBroker;
			_quickInfoBroker = quickInfoBroker;
		}

		internal DisplayWindowHelper ForTextView(ITextView view)
		{
			return new DisplayWindowHelper(
				view,
				_completionBroker,
				_signatureHelpBroker,
				_smartTagBroker,
				_quickInfoBroker);
		}

		internal ITextView TextView { get; private set; }

		internal bool IsCompletionActive
		{
			get { return _completionBroker != null ? _completionBroker.IsCompletionActive(this.TextView) : false; }
		}

		internal bool IsSignatureHelpActive
		{
			get { return _signatureHelpBroker != null ? _signatureHelpBroker.IsSignatureHelpActive(this.TextView) : false; }
		}

		internal bool IsSmartTagSessionActive
		{
			get { return _smartTagBroker != null ? _smartTagBroker.IsLightBulbSessionActive(this.TextView) : false; }
		}

		internal bool IsQuickInfoActive
		{
			get { return _quickInfoBroker != null ? _quickInfoBroker.IsQuickInfoActive(this.TextView) : false; }
		}
	}
}
