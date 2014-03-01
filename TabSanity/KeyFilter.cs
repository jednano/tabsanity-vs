using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Formatting;
using System;
using IServiceProvider = System.IServiceProvider;

namespace TabSanity
{
	internal abstract class KeyFilter : TabOptionsListener, IOleCommandTarget
	{
		internal bool Added = false;
		internal IOleCommandTarget NextTarget = null;
		protected DisplayWindowHelper DisplayHelper = null;

		private IServiceProvider _serviceProvider = null;

		#region Computed Properties

		protected ITextCaret Caret
		{
			get { return TextView.Caret; }
		}

		protected bool CaretCharIsASpace
		{
			get
			{
				return Caret.Position.BufferPosition.Position < Caret.Position.BufferPosition.Snapshot.Length
					&& Caret.Position.BufferPosition.GetChar() == ' ';
			}
		}

		protected int CaretColumn
		{
			get { return Caret.Position.BufferPosition.Position - CaretLine.Start.Position; }
		}

		protected bool CaretIsWithinCodeRange
		{
			get { return CaretColumn > ColumnAfterLeadingSpaces; }
		}

		protected ITextViewLine CaretLine
		{
			get { return Caret.ContainingTextViewLine; }
		}

		protected bool CaretPrevCharIsASpace
		{
			get
			{
				return Caret.Position.BufferPosition.Position > 0
					&& Caret.Position.BufferPosition.Subtract(1).GetChar() == ' ';
			}
		}

		protected int ColumnAfterLeadingSpaces
		{
			get
			{
				var snapshot = CaretLine.Snapshot;
				var column = 0;
				for (var i = CaretLine.Start.Position; i < CaretLine.End.Position; i++)
				{
					column++;
					if (snapshot[i] != ' ') break;
				}
				return column;
			}
		}

		protected int ColumnBeforeTrailingSpaces
		{
			get
			{
				var snapshot = CaretLine.Snapshot;
				var column = CaretLine.Length;
				for (var i = CaretLine.End.Position - 1; i > CaretLine.Start.Position; i--)
				{
					column--;
					if (snapshot[i] != ' ') break;
				}
				return column;
			}
		}

		protected bool IsInAutomationFunction
		{
			get { return VsShellUtilities.IsInAutomationFunction(_serviceProvider); }
		}

		protected int VirtualCaretColumn
		{
			get
			{
				return Caret.Position.BufferPosition.Position +
					   Caret.Position.VirtualBufferPosition.VirtualSpaces - CaretLine.Start.Position;
			}
		}

		#endregion Computed Properties

		public KeyFilter(DisplayWindowHelper displayHelper, IWpfTextView textView, IServiceProvider provider)
			: base(textView)
		{
			DisplayHelper = displayHelper;
			_serviceProvider = provider;
		}

		public abstract int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut);

		public abstract int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText);
	}
}
