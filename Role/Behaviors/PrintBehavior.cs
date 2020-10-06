using Microsoft.Xaml.Behaviors;

using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace RoleDDNG.Role.Behaviors
{
    public class PrintBehavior : Behavior<Button>
    {
        public TextBox TextSource
        {
            get { return (TextBox)GetValue(TextSourceProperty); }
            set { SetValue(TextSourceProperty, value); }
        }

        public static readonly DependencyProperty TextSourceProperty =
            DependencyProperty.Register(nameof(TextSource), typeof(TextBox), typeof(PrintBehavior), new PropertyMetadata(null));

        protected override void OnAttached()
        {
            base.OnAttached();
            AssociatedObject.AddHandler(Button.ClickEvent, new RoutedEventHandler(AssociatedObject_Click), true);
        }

        private void AssociatedObject_Click(object sender, RoutedEventArgs e)
        {
            var textBox = GetValue(TextSourceProperty);

            if (!(textBox is Visual))
            {
                return;
            }
            PrintDialog printDialog = new PrintDialog();
            FlowDocument printedDocument = new FlowDocument();
            Section section = new Section();
            Paragraph paragraph = new Paragraph();
            paragraph.Inlines.Add(TextSource.Text);
            section.Blocks.Add(paragraph);
            printedDocument.Blocks.Add(section);
            printDialog.PrintDocument(((IDocumentPaginatorSource)printedDocument).DocumentPaginator, "");
        }

        protected override void OnDetaching()
        {
            base.OnDetaching();
            AssociatedObject.RemoveHandler(Button.ClickEvent, new RoutedEventHandler(AssociatedObject_Click));
        }
    }
}