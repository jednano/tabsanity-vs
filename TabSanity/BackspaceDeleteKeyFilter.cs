using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using System;
using IServiceProvider = System.IServiceProvider;

namespace TabSanity
{
	internal class BackspaceDeleteKeyFilter : KeyFilter
	{
		public BackspaceDeleteKeyFilter(DisplayWindowHelper displayHelper, IWpfTextView textView, IServiceProvider provider)
			: base(displayHelper, textView, provider)
		{
		}

		public override int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
		{
			if (ConvertTabsToSpaces
				&& TextView.Selection.IsEmpty
				&& !CaretIsWithinCodeRange
				&& !IsInAutomationFunction
				&& !DisplayHelper.IsCompletionActive
				&& !DisplayHelper.IsSignatureHelpActive
				)
			{
				var handled = false;

				if (pguidCmdGroup == VSConstants.VSStd2K)
				{
					switch ((VSConstants.VSStd2KCmdID)nCmdID)
					{
						case VSConstants.VSStd2KCmdID.BACKSPACE:
							handled = HandleBackspaceKey();
							break;

						case VSConstants.VSStd2KCmdID.DELETE:
							handled = HandleDeleteKey();
							break;
					}
				}
				else if (pguidCmdGroup == VSConstants.GUID_VSStandardCommandSet97)
				{
					switch ((VSConstants.VSStd97CmdID)nCmdID)
					{
						case VSConstants.VSStd97CmdID.Delete:
							handled = HandleDeleteKey();
							break;
					}
				}

				if (handled)
				{
					return VSConstants.S_OK;
				}
			}

			return NextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
		}

		public override int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
		{
			return NextTarget.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
		}

		private bool HandleBackspaceKey()
		{
			ReplaceVirtualSpaces();

			var snapshot = TextView.TextBuffer.CurrentSnapshot;
			var caretPos = Caret.Position.BufferPosition.Position;

			// Determine the number of spaces until the previous tab stop.
			var spacesToRemove = ((CaretColumn - 1) % IndentSize) + 1;

			// Make sure we only delete spaces.
			for (var i = 0; i < spacesToRemove; i++)
			{
				var snapshotPos = caretPos - 1 - i;
				if (snapshotPos < 0 || snapshot[snapshotPos] != ' ')
				{
					spacesToRemove = i;
					break;
				}
			}

			if (spacesToRemove > 1)
			{
				TextView.TextBuffer.Delete(new Span(caretPos - spacesToRemove, spacesToRemove));
				return true;
			}
			else
			{
				return false;
			}
		}

		private bool HandleDeleteKey()
		{
			// If we are in virtual space, we should already be at the end of the line,
			// so let Visual Studio handle the keypress.
			if (Caret.InVirtualSpace)
			{
				return false;
			}

			var snapshot = TextView.TextBuffer.CurrentSnapshot;
			var caretPos = Caret.Position.BufferPosition.Position;

			// Determine the number of spaces until the next tab stop.
			var spacesToRemove = IndentSize - (CaretColumn % IndentSize);

			// Make sure we only delete spaces.
			for (var i = 0; i < spacesToRemove; i++)
			{
				var snapshotPos = caretPos + i;
				if (snapshotPos >= snapshot.Length || snapshot[snapshotPos] != ' ')
				{
					spacesToRemove = i;
					break;
				}
			}

			if (spacesToRemove > 1)
			{
				TextView.TextBuffer.Delete(new Span(caretPos, spacesToRemove));
				return true;
			}
			else
			{
				return false;
			}
		}

		private void ReplaceVirtualSpaces()
		{
			if (Caret.InVirtualSpace)
			{
				TextView.TextBuffer.Insert(Caret.Position.BufferPosition, new string(' ', Caret.Position.VirtualSpaces));
				Caret.MoveTo(CaretLine.End);
			}
		}
	}
}
