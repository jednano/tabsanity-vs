using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;
using System.ComponentModel.Composition;

namespace TabSanity
{
	[Export(typeof (IVsTextViewCreationListener))]
	[ContentType("text")]
	[TextViewRole(PredefinedTextViewRoles.Editable)]
	internal class KeyFilterFactory : IVsTextViewCreationListener
	{
		[Import(typeof (IVsEditorAdaptersFactoryService))] private IVsEditorAdaptersFactoryService _editorFactory;

		[Import] private SVsServiceProvider _serviceProvider;

		public void VsTextViewCreated(IVsTextView viewAdapter)
		{
			var app = (DTE) _serviceProvider.GetService(typeof (DTE));
			var view = _editorFactory.GetWpfTextView(viewAdapter);
			if (app == null || view == null)
				return;

			// ReSharper disable ObjectCreationAsStatement
			new BackspaceDeleteKeyFilter(app, view);
			// ReSharper restore ObjectCreationAsStatement

			AddCommandFilter(viewAdapter, new ArrowKeyFilter(view, viewAdapter));
		}

		private static void AddCommandFilter(IVsTextView viewAdapter, ArrowKeyFilter commandFilter)
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
