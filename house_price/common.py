# house_price/common.py
# 模式核心配置（所有页面共享）
CHART_MODE_CONFIGS = {
    "simple": {
        "trend_month": 6,       # 时间范围：6个月
        "chart_type": "line",   # 图表类型：折线
        "show_legend": False,   # 隐藏图例
        "show_label": False     # 隐藏数值标签
    },
    "detailed": {
        "trend_month": 24,      # 时间范围：24个月
        "chart_type": "bar",    # 图表类型：柱状
        "show_legend": True,    # 显示图例
        "show_label": True      # 显示数值标签
    }
}

# 通用工具函数：获取合法的模式配置（避免非法参数）
def get_valid_chart_config(mode):
    valid_modes = ["simple", "detailed"]
    if mode not in valid_modes:
        mode = "detailed"  # 默认详细模式
    return CHART_MODE_CONFIGS[mode], mode