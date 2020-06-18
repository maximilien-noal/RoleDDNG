using RoleDDNG.Role.UserControls.Traits;
using System.Windows.Controls;

namespace RoleDDNG.Role.UserControls
{
    /// <summary>
    /// Interaction logic for TownGeneratorUserControl.xaml.
    /// </summary>
    public partial class TownGeneratorUserControl : UserControl, IUpdateMdiWindowIconTrait
    {
        public TownGeneratorUserControl()
        {
            this.InitializeComponent();
            this.Loaded += (s, e) => ((IUpdateMdiWindowIconTrait)this).ChangeMdiWindowIcon(this);
        }
    }
}