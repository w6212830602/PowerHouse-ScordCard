using CommunityToolkit.Mvvm.ComponentModel;

namespace ScoreCard.ViewModels
{
    /// <summary>
    /// 用於跟踪銷售代表選擇狀態的類
    /// </summary>
    public partial class RepSelectionItem : ObservableObject
    {
        /// <summary>
        /// 銷售代表姓名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 是否被選中
        /// </summary>
        [ObservableProperty]
        private bool _isSelected;
    }
}
