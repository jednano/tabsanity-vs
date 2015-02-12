using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using System;
using IServiceProvider = System.IServiceProvider;

namespace TabSanity
{
	internal class ArrowKeyFilter : KeyFilter
	{
		private int? _savedCaretColumn;
		private ITextSnapshotLine _snapshotLine;

		#region Arrow key constants

		private const uint ARROW_LEFT = (uint)VSConstants.VSStd2KCmdID.LEFT;
		private const uint ARROW_RIGHT = (uint)VSConstants.VSStd2KCmdID.RIGHT;
		private const uint ARROW_UP = (uint)VSConstants.VSStd2KCmdID.UP;
		private const uint ARROW_DOWN = (uint)VSConstants.VSStd2KCmdID.DOWN;

		#endregion Arrow key constants

		public ArrowKeyFilter(DisplayWindowHelper displayHelper, IWpfTextView textView, IServiceProvider provider)
			: base(displayHelper, textView, provider)
		{
			Caret.PositionChanged += CaretOnPositionChanged;
		}

		private void CaretOnPositionChanged(object sender, CaretPositionChangedEventArgs e)
		{
			_savedCaretColumn = null;
		}

		public override int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
		{
			return NextTarget.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
		}

		public override int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
		{
			if (ConvertTabsToSpaces
				&& TextView.Selection.IsEmpty
				&& pguidCmdGroup == VSConstants.VSStd2K
				&& !IsInAutomationFunction
				&& !DisplayHelper.IsCompletionActive
				&& !DisplayHelper.IsSignatureHelpActive
				)
			{
				switch (nCmdID)
				{
					case ARROW_LEFT:
					case ARROW_RIGHT:
						if (CaretIsWithinCodeRange)
							goto default;
						break;

					case ARROW_UP:
					case ARROW_DOWN:
						Caret.PositionChanged -= CaretOnPositionChanged;
						if (!_savedCaretColumn.HasValue)
							_savedCaretColumn = VirtualCaretColumn;
						break;

					default:
						_savedCaretColumn = null;
						return NextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
				}

				switch (nCmdID)
				{
					case ARROW_LEFT:
						Caret.MoveToPreviousCaretPosition();
						if (CaretCharIsASpace)
						{
							MoveCaretToPreviousTabStop();
							return VSConstants.S_OK;
						}
						Caret.MoveToNextCaretPosition();
						break;

					case ARROW_RIGHT:
						if (CaretCharIsASpace)
						{
							Caret.MoveToNextCaretPosition();
							MoveCaretToNextTabStop();
							return VSConstants.S_OK;
						}
						break;

					case ARROW_UP:
						try
						{
							_snapshotLine = TextView.TextSnapshot.GetLineFromPosition(
								Caret.Position.BufferPosition.Subtract(CaretColumn + 1));
							MoveCaretToNearestVirtualTabStop();
						}
						catch (ArgumentOutOfRangeException)
						{
						}
						Caret.PositionChanged += CaretOnPositionChanged;
						return VSConstants.S_OK;

					case ARROW_DOWN:
						try
						{
							_snapshotLine = FindNextLine();
							MoveCaretToNearestVirtualTabStop();
						}
						catch (ArgumentOutOfRangeException)
						{
						}
						Caret.PositionChanged += CaretOnPositionChanged;
						return VSConstants.S_OK;
				}
			}

			return NextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
		}

		ITextSnapshotLine FindNextLine()
		{
			var snapshot = TextView.TextSnapshot;
			var caretBufferPosition = Caret.Position.BufferPosition;
			var current = snapshot.GetLineFromPosition(caretBufferPosition.Position);
			var next = snapshot.GetLineFromPosition(caretBufferPosition.Add(CaretLine.Length - CaretColumn + 2));
			if (next.LineNumber == current.LineNumber && next.LineNumber + 1 < snapshot.LineCount)
				next = snapshot.GetLineFromLineNumber(next.LineNumber + 1);
			return next;
		}

		private void MoveCaretToNearestVirtualTabStop()
		{
			if (!_savedCaretColumn.HasValue)
				return;

			var lastIndentColumn = ColumnAfterLeadingSpaces - 1;
			if (lastIndentColumn < 0) lastIndentColumn = 0;
			MoveCaretToVirtualPosition(_savedCaretColumn.Value);

			if (Caret.InVirtualSpace)
			{
				var eol = CaretLine.Length;
				MoveCaretToVirtualPosition((eol > 0) ? eol : lastIndentColumn);
				// TODO: lastIndentColumn offset?
				return;
			}

			if (CaretColumn > ColumnAfterLeadingSpaces)
				return;

			var remainder = CaretColumn % IndentSize;
			if (remainder == 0)
				return;

			if (remainder < IndentSize / 2)
				MoveCaretToPreviousTabStop();
			else
				MoveCaretToNextTabStop();
		}

		private void MoveCaretToVirtualPosition(int pos)
		{
			Caret.MoveTo(new VirtualSnapshotPoint(_snapshotLine, pos));
			Caret.EnsureVisible();
		}

		private void MoveCaretToNextTabStop()
		{
			while (CaretColumn % IndentSize != 0 && CaretCharIsASpace)
				Caret.MoveToNextCaretPosition();
			Caret.EnsureVisible();
		}

		private void MoveCaretToPreviousTabStop()
		{
			Caret.MoveToPreviousCaretPosition();
			while (CaretColumn % IndentSize != (IndentSize - 1) && CaretCharIsASpace)
				Caret.MoveToPreviousCaretPosition();
			Caret.MoveToNextCaretPosition();
			Caret.EnsureVisible();
		}

		public override void Dispose()
		{
			Caret.PositionChanged -= CaretOnPositionChanged;
			base.Dispose();
		}
	}
}
