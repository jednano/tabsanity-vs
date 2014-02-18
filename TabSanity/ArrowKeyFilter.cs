using System;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Formatting;

namespace TabSanity
{
	internal class ArrowKeyFilter : TabOptionsListener, IOleCommandTarget
	{
		internal IOleCommandTarget NextTarget;
		internal bool Added;
		private int? _savedCaretColumn;
		private ITextSnapshotLine _snapshotLine;
		private ITextView _view;
		private ICompletionBroker _broker;

		#region Arrow key constants

		private const uint ARROW_LEFT = (uint) VSConstants.VSStd2KCmdID.LEFT;
		private const uint ARROW_RIGHT = (uint) VSConstants.VSStd2KCmdID.RIGHT;
		private const uint ARROW_UP = (uint) VSConstants.VSStd2KCmdID.UP;
		private const uint ARROW_DOWN = (uint) VSConstants.VSStd2KCmdID.DOWN;

		#endregion

		#region Computed Properties

		private ITextCaret Caret
		{
			get { return TextView.Caret; }
		}

		private ITextViewLine CaretLine
		{
			get { return Caret.ContainingTextViewLine; }
		}

		private int CaretColumn
		{
			get { return Caret.Position.BufferPosition.Position - CaretLine.Start.Position; }
		}

		private int VirtualCaretColumn
		{
			get
			{
				return Caret.Position.BufferPosition.Position +
					   Caret.Position.VirtualBufferPosition.VirtualSpaces - CaretLine.Start.Position;
			}
		}

		private bool CaretCharIsASpace
		{
			get { return Caret.Position.BufferPosition.GetChar() == ' '; }
		}

		private int ColumnAfterLeadingSpaces
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

		private int ColumnBeforeTrailingSpaces
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

		private bool CaretIsWithinCodeRange
		{
			get
			{
				return CaretColumn > ColumnAfterLeadingSpaces &&
					   CaretColumn < ColumnBeforeTrailingSpaces;
			}
		}

		#endregion

		public ArrowKeyFilter(ICompletionBroker broker, IWpfTextView textView)
			: base(textView)
		{
			_broker = broker;
			_view = textView;
			Caret.PositionChanged += CaretOnPositionChanged;
		}

		private void CaretOnPositionChanged(object sender, CaretPositionChangedEventArgs e)
		{
			_savedCaretColumn = null;
		}

		public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
		{
			return NextTarget.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
		}

		public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
		{
			if (ConvertTabsToSpaces 
				&& TextView.Selection.IsEmpty 
				&& pguidCmdGroup == VSConstants.VSStd2K 
				&& (_broker == null || !_broker.IsCompletionActive(_view))
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
							_snapshotLine = TextView.TextSnapshot.GetLineFromPosition(
								Caret.Position.BufferPosition.Add(CaretLine.Length - CaretColumn + 2));
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

			var remainder = CaretColumn%IndentSize;
			if (remainder == 0)
				return;

			if (remainder < IndentSize/2)
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
			while (CaretColumn%IndentSize != 0 && CaretCharIsASpace)
				Caret.MoveToNextCaretPosition();
			Caret.EnsureVisible();
		}

		private void MoveCaretToPreviousTabStop()
		{
			Caret.MoveToPreviousCaretPosition();
			while (CaretColumn%IndentSize != (IndentSize - 1) && CaretCharIsASpace)
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
