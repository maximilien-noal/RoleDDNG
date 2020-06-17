using RoleDDNG.Role.UserControls.Traits;
using System.Windows.Controls;

namespace RoleDDNG.Role.UserControls
{
    /// <summary>
    /// Interaction logic for OpenCharacterUserControl.xaml
    /// </summary>
    public partial class OpenCharacterUserControl : UserControl, IUpdateMdiWindowIconTrait
    {
        public OpenCharacterUserControl()
        {
            InitializeComponent();
            this.Loaded += (s, e) => ((IUpdateMdiWindowIconTrait)this).ChangeMdiWindowIcon(this);
        }
    }
}