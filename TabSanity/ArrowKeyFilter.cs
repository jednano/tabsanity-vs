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
		private object _activeLock = new object();
		private bool _allowClearSavedCaretColumn = true;
		private bool _isActive;

		#region Arrow key constants
		private const uint ARROW_LEFT = (uint)VSConstants.VSStd2KCmdID.LEFT;
		private const uint SHIFT_ARROW_LEFT = (uint)VSConstants.VSStd2KCmdID.LEFT_EXT;
		private const uint ARROW_RIGHT = (uint)VSConstants.VSStd2KCmdID.RIGHT;
		private const uint SHIFT_ARROW_RIGHT = (uint)VSConstants.VSStd2KCmdID.RIGHT_EXT;
		private const uint ARROW_UP = (uint)VSConstants.VSStd2KCmdID.UP;
		private const uint SHIFT_ARROW_UP = (uint)VSConstants.VSStd2KCmdID.UP_EXT;
		private const uint ARROW_DOWN = (uint)VSConstants.VSStd2KCmdID.DOWN;
		private const uint SHIFT_ARROW_DOWN = (uint)VSConstants.VSStd2KCmdID.DOWN_EXT;
		#endregion Arrow key constants

		public ArrowKeyFilter(DisplayWindowHelper displayHelper, IWpfTextView textView, IServiceProvider provider)
			: base(displayHelper, textView, provider)
		{
			Caret.PositionChanged += CaretOnPositionChanged;
		}

		private void CaretOnPositionChanged(object sender, CaretPositionChangedEventArgs e)
		{
			if (_allowClearSavedCaretColumn)
				_savedCaretColumn = null;

			if (_isActive
				|| IsInAutomationFunction
				|| DisplayHelper.IsCompletionActive
				|| DisplayHelper.IsSignatureHelpActive
				|| CaretIsWithinCodeRange
				|| CaretColumn % IndentSize == 0)
				return;

			lock (_activeLock)
			{
				var originalSavedCaretColumn = _savedCaretColumn;
				var originalSnapshotLine = _snapshotLine;
				try
				{
					_isActive = true;
					_savedCaretColumn = CaretColumn;
					_snapshotLine = TextView.TextSnapshot.GetLineFromPosition(Caret.Position.BufferPosition);
					var caretStartingPosition = Caret.Position.VirtualBufferPosition;

					MoveCaretToNearestVirtualTabStop();

					if (!TextView.Selection.IsEmpty)
						AdjustTextSelection(0, TextView.Selection.AnchorPoint);
				}
				finally
				{
					_isActive = false;
					_savedCaretColumn = originalSavedCaretColumn;
					_snapshotLine = originalSnapshotLine;
				}
			}
		}

		public override int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
		{
			return NextTarget.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
		}

		public override int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
		{
			if (nCmdID < ARROW_LEFT
				|| nCmdID > SHIFT_ARROW_DOWN
				|| pguidCmdGroup != VSConstants.VSStd2K
				|| IsInAutomationFunction
				|| DisplayHelper.IsCompletionActive
				|| DisplayHelper.IsSignatureHelpActive)
			{
				_savedCaretColumn = null;
				return NextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
			}

			lock (_activeLock)
			{
				try
				{
					_isActive = true;

					switch (nCmdID)
					{
						case ARROW_LEFT:
						case ARROW_RIGHT:
						case SHIFT_ARROW_LEFT:
						case SHIFT_ARROW_RIGHT:
							if (CaretIsWithinCodeRange)
								goto default;
							break;

						case ARROW_UP:
						case ARROW_DOWN:
						case SHIFT_ARROW_UP:
						case SHIFT_ARROW_DOWN:
							_allowClearSavedCaretColumn = false;
							if (!_savedCaretColumn.HasValue)
								_savedCaretColumn = VirtualCaretColumn;
							break;

						default:
							_savedCaretColumn = null;
							return NextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
					}

					var caretStartingPosition = Caret.Position.VirtualBufferPosition;
					var selectionAnchorPoint = TextView.Selection.IsEmpty ? caretStartingPosition : TextView.Selection.AnchorPoint;
					switch (nCmdID)
					{
						case ARROW_LEFT:
						case SHIFT_ARROW_LEFT:
							Caret.MoveToPreviousCaretPosition();
							if (CaretCharIsASpace)
							{
								MoveCaretToPreviousTabStop();
								AdjustTextSelection(nCmdID, selectionAnchorPoint);
								return VSConstants.S_OK;
							}
							Caret.MoveToNextCaretPosition();
							break;

						case ARROW_RIGHT:
						case SHIFT_ARROW_RIGHT:
							if (CaretCharIsASpace)
							{
								Caret.MoveToNextCaretPosition();
								MoveCaretToNextTabStop();
								AdjustTextSelection(nCmdID, selectionAnchorPoint);
								return VSConstants.S_OK;
							}
							break;

						case ARROW_UP:
						case SHIFT_ARROW_UP:
							try
							{
								_snapshotLine = TextView.TextSnapshot.GetLineFromPosition(
									Caret.Position.BufferPosition.Subtract(CaretColumn + 1));
								MoveCaretToNearestVirtualTabStop();
							}
							catch (ArgumentOutOfRangeException)
							{
							}
							_allowClearSavedCaretColumn = true;
							AdjustTextSelection(nCmdID, selectionAnchorPoint);
							return VSConstants.S_OK;

						case ARROW_DOWN:
						case SHIFT_ARROW_DOWN:
							try
							{
								_snapshotLine = FindNextLine();
								MoveCaretToNearestVirtualTabStop();
							}
							catch (ArgumentOutOfRangeException)
							{
							}
							_allowClearSavedCaretColumn = true;
							AdjustTextSelection(nCmdID, selectionAnchorPoint);
							return VSConstants.S_OK;
					}

					return NextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
				}
				finally
				{
					_isActive = false;
				}
			}
		}

		private void AdjustTextSelection(uint nCmdID, VirtualSnapshotPoint selectionAnchorPoint)
		{
			if (nCmdID % 2 == 0) // nCmdID is even for all shift arrows
			{
				// Adjust the text selection if shift is being held
				TextView.Selection.Select(selectionAnchorPoint, Caret.Position.VirtualBufferPosition);
			}
			else if (!TextView.Selection.IsEmpty)
			{
				// Turn off selection if a regular arrow was pressed
				TextView.Selection.Select(TextView.Selection.AnchorPoint, TextView.Selection.AnchorPoint);
			}
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
			var caretStartingPosition = Caret.Position.VirtualBufferPosition;
			while (CaretColumn % IndentSize != 0 && CaretCharIsASpace)
				Caret.MoveToNextCaretPosition();

			if (CaretColumn % IndentSize != 0)
				Caret.MoveTo(caretStartingPosition); // Do not align on non-exact tab stops

			Caret.EnsureVisible();
		}

		private void MoveCaretToPreviousTabStop()
		{
			var caretStartingPosition = Caret.Position.VirtualBufferPosition;
			var caretStartingColumn = CaretColumn;
			Caret.MoveToPreviousCaretPosition();
			var lastCaretColumn = -1;
			while (CaretColumn % IndentSize != (IndentSize - 1) && CaretCharIsASpace)
			{
				if (CaretColumn == lastCaretColumn)
					break; // Prevent infinite loop on first char of first line

				lastCaretColumn = CaretColumn;
				Caret.MoveToPreviousCaretPosition();
			}

			if (Caret.Position.BufferPosition.Position != 0)
			{ // Do this for all cases except the first char of the document
				Caret.MoveToNextCaretPosition();
			}

			VirtualSnapshotPoint? caretNewPosition = Caret.Position.VirtualBufferPosition;
			int movedBy = caretStartingColumn - CaretColumn;
			if (movedBy % IndentSize != 0)
			{
				// We moved less than a full tab stop length. Only allow this if the cursor started in the middle of a full tab
				for (int i = 0; i < IndentSize; i++)
				{
					if (!CaretCharIsASpace)
					{
						caretNewPosition = null;
						Caret.MoveTo(caretStartingPosition); // Do not align on non-exact tab stops
						break;
					}
					Caret.MoveToNextCaretPosition();
				}
				if (caretNewPosition != null)
				{
					Caret.MoveTo(caretNewPosition.Value); // Go back to original new position
				}
			}

			Caret.EnsureVisible();
		}

		public override void Dispose()
		{
			Caret.PositionChanged -= CaretOnPositionChanged;
			base.Dispose();
		}
	}
}
