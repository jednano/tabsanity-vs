using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Text.Editor;

namespace TabSanity
{
	internal class BackspaceDeleteKeyFilter : TabOptionsListener
	{
		private readonly TextDocumentKeyPressEvents _keyPressEvents;
		private const string BACKSPACE_KEY = "\b";
		private const string DELETE_KEY = "";

		public BackspaceDeleteKeyFilter(_DTE app, IWpfTextView textView)
			: base(textView)
		{
			var events = (Events2)app.Events;
			// ReSharper disable RedundantArgumentDefaultValue
			_keyPressEvents = events.TextDocumentKeyPressEvents[null];
			// ReSharper restore RedundantArgumentDefaultValue
			_keyPressEvents.BeforeKeyPress += BeforeKeyPress;
		}

		protected override void OnConvertTabsToSpacesOptionChanged()
		{
			base.OnConvertTabsToSpacesOptionChanged();
			_keyPressEvents.BeforeKeyPress -= BeforeKeyPress;
			if (ConvertTabsToSpaces)
				_keyPressEvents.BeforeKeyPress += BeforeKeyPress;
		}

		private void BeforeKeyPress(string keypress, TextSelection selection,
									bool inStatementCompletion, ref bool cancelKeypress)
		{
			if (!ConvertTabsToSpaces || !selection.IsEmpty || inStatementCompletion) return;

			switch (keypress)
			{
				case BACKSPACE_KEY:
					ReplaceVirtualSpaces(selection);

					do
					{
						selection.CharLeft(true);
						if (selection.Text == " ")
						{
							cancelKeypress = true;
							selection.Delete();
							continue;
						}
						selection.CharRight(true);
						return;
					} while (selection.CurrentColumn % IndentSize != 1);
					return;

				case DELETE_KEY:
					ReplaceVirtualSpaces(selection);

					for (var i = 0; i < IndentSize; i++)
					{
						selection.CharRight(true);
						if (selection.Text == " ")
						{
							cancelKeypress = true;
							selection.Delete();
							continue;
						}
						selection.CharLeft(true);
						return;
					}
					return;
			}
		}

		private void ReplaceVirtualSpaces(TextSelection selection)
		{
			if (TextView.Caret.InVirtualSpace)
			{
				var spaces = TextView.Caret.Position.VirtualSpaces;
				selection.DeleteWhitespace();
				selection.PadToColumn(spaces + 1);
			}
		}

		public override void Dispose()
		{
			_keyPressEvents.BeforeKeyPress -= BeforeKeyPress;
			base.Dispose();
		}
	}
}
