using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;
using System.ComponentModel.Composition;

namespace TabSanity
{
	[Export(typeof(IVsTextViewCreationListener))]
	[ContentType("text")]
	[TextViewRole(PredefinedTextViewRoles.Editable)]
	internal class KeyFilterFactory : IVsTextViewCreationListener
	{
		[Import(typeof(IVsEditorAdaptersFactoryService))]
		private IVsEditorAdaptersFactoryService _editorFactory;

		[Import]
		private SVsServiceProvider _serviceProvider;

		private DisplayWindowHelper _helperFactory;

		[ImportingConstructor]
		internal KeyFilterFactory(
			ICompletionBroker completionBroker,
			ISignatureHelpBroker signatureHelpBroker,
			ISmartTagBroker smartTagBroker,
			IQuickInfoBroker quickInfoBroker)
		{
			_helperFactory = new DisplayWindowHelper(completionBroker, signatureHelpBroker, smartTagBroker, quickInfoBroker);
		}

		public void VsTextViewCreated(IVsTextView viewAdapter)
		{
			var view = _editorFactory.GetWpfTextView(viewAdapter);
			if (view == null)
				return;

			var displayHelper = _helperFactory.ForTextView(view);

			AddCommandFilter(viewAdapter, new ArrowKeyFilter(displayHelper, view, _serviceProvider));
			AddCommandFilter(viewAdapter, new BackspaceDeleteKeyFilter(displayHelper, view, _serviceProvider));
		}

		private static void AddCommandFilter(IVsTextView viewAdapter, KeyFilter commandFilter)
		{
			if (commandFilter.Added) return;
			//get the view adapter from the editor factory
			IOleCommandTarget next;
			var hr = viewAdapter.AddCommandFilter(commandFilter, out next);

			if (hr != VSConstants.S_OK) return;
			commandFilter.Added = true;
			//you'll need the next target for Exec and QueryStatus
			if (next != null)
				commandFilter.NextTarget = next;
		}
	}
}
