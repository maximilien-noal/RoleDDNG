using System.Windows.Controls;

using RoleDDNG.Role.UserControls.Traits;

namespace RoleDDNG.Role.UserControls
{
    /// <summary>
    /// Interaction logic for CharactersXpUserControl.xaml
    /// </summary>
    public partial class CharactersXpUserControl : UserControl, IUpdateMdiWindowIconTrait
    {
        public CharactersXpUserControl()
        {
            InitializeComponent();
            this.Loaded += (s, e) => ((IUpdateMdiWindowIconTrait)this).ChangeMdiWindowIcon(this);
        }
    }
}