using RoleDDNG.Role.UserControls.Traits;
using System.Windows.Controls;

namespace RoleDDNG.Role.UserControls
{
    /// <summary>
    /// Interaction logic for DiceRollUserControl.xaml
    /// </summary>
    public partial class DiceRollUserControl : UserControl, IUpdateMdiWindowIconTrait
    {
        public DiceRollUserControl()
        {
            InitializeComponent();
            this.Loaded += (s, e) => ((IUpdateMdiWindowIconTrait)this).ChangeMdiWindowIcon(this);
        }
    }
}