# -*- coding: utf-8 -*-
"""
Excel折线图绘制工具 (Web版本)
基于Streamlit的现代化Web界面
"""

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
from io import BytesIO

# 配置matplotlib中文支持
matplotlib.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'SimHei', 'STHeiti', 'DejaVu Sans']
matplotlib.rcParams['axes.unicode_minus'] = False

# 页面配置
st.set_page_config(
    page_title="Excel折线图绘制工具",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS样式
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        background-color: #d4edda;
        border-left: 4px solid #28a745;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # 页面标题
    st.markdown('<div class="main-header">📊 Excel折线图绘制工具</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">专为科研人员设计的专业绘图工具</div>', unsafe_allow_html=True)

    # 侧边栏 - 文件上传和数据选择
    with st.sidebar:
        st.header("📁 数据加载")

        uploaded_file = st.file_uploader(
            "选择Excel文件",
            type=['xlsx', 'xls'],
            help="支持 .xlsx 和 .xls 格式"
        )

        if uploaded_file:
            try:
                df = pd.read_excel(uploaded_file)
                st.success(f"✅ 成功加载 {len(df)} 行数据")

                # 显示数据预览
                with st.expander("📋 数据预览", expanded=False):
                    st.dataframe(df.head(10))

                st.divider()

                # 列选择
                st.header("🎯 数据选择")

                columns = df.columns.tolist()

                # X轴选择
                x_col = st.selectbox(
                    "横坐标 (X轴)",
                    columns,
                    help="选择作为横坐标的列"
                )

                # Y轴选择
                default_y = [col for col in columns if col != x_col]
                y_cols = st.multiselect(
                    "纵坐标 (Y轴，可多选)",
                    columns,
                    default=default_y,
                    help="可以选择多个列在同一图表中显示"
                )

                if not y_cols:
                    st.warning("⚠️ 请至少选择一个Y轴列")
                    return

            except Exception as e:
                st.error(f"❌ 读取文件失败: {str(e)}")
                return
        else:
            st.info("👆 请上传Excel文件开始")
            return

    # 主内容区
    if uploaded_file and y_cols:
        # 创建两列布局
        col1, col2 = st.columns([1, 2])

        with col1:
            st.header("🎨 样式设置")

            # 标题和标签
            with st.expander("📝 标题和标签", expanded=True):
                title = st.text_input("图表标题", value="Excel数据折线图")
                xlabel = st.text_input("X轴标签", value=x_col)
                ylabel = st.text_input("Y轴标签", value="Value")

            # 字体设置
            with st.expander("🔤 字体大小", expanded=True):
                title_fontsize = st.slider("标题字号", 8, 30, 14)
                label_fontsize = st.slider("坐标轴字号", 8, 24, 12)
                legend_fontsize = st.slider("图例字号", 6, 20, 10)

            # 线条样式
            with st.expander("📏 线条样式", expanded=True):
                linewidth = st.slider("线宽", 0.5, 10.0, 2.0, 0.5)
                markersize = st.slider("标记点大小", 0, 20, 4)

                linestyle = st.selectbox(
                    "线型",
                    ['-', '--', '-.', ':'],
                    format_func=lambda x: {
                        '-': '实线 (－)',
                        '--': '虚线 (- -)',
                        '-.': '点划线 (-.-.)',
                        ':': '点线 (···)'
                    }[x]
                )

                marker = st.selectbox(
                    "标记样式",
                    ['无', 'o', 's', '^', 'v', 'D', '*', '+', 'x'],
                    format_func=lambda x: {
                        '无': '无标记',
                        'o': '圆圈 ●',
                        's': '方块 ■',
                        '^': '上三角 ▲',
                        'v': '下三角 ▼',
                        'D': '菱形 ◆',
                        '*': '星号 ✱',
                        '+': '加号 +',
                        'x': '叉号 ×'
                    }[x]
                )

            # 网格和背景
            with st.expander("🎭 网格和背景", expanded=True):
                show_grid = st.checkbox("显示网格", value=True)
                grid_alpha = st.slider("网格透明度", 0.0, 1.0, 0.3, 0.1)

                bg_color = st.color_picker("背景颜色", value="#FFFFFF")

            # 坐标轴范围
            with st.expander("📐 坐标轴范围", expanded=False):
                auto_axis = st.checkbox("自动范围", value=True)

                if not auto_axis:
                    col_a, col_b = st.columns(2)
                    with col_a:
                        xmin = st.number_input("X最小值", value=float(df[x_col].min()))
                        ymin = st.number_input("Y最小值", value=0.0)
                    with col_b:
                        xmax = st.number_input("X最大值", value=float(df[x_col].max()))
                        ymax = st.number_input("Y最大值", value=1.0)

            # 图表尺寸
            with st.expander("📏 图表尺寸", expanded=False):
                fig_width = st.slider("宽度 (英寸)", 6, 20, 12)
                fig_height = st.slider("高度 (英寸)", 4, 15, 6)
                dpi = st.selectbox("分辨率 (DPI)", [100, 150, 200, 300, 600], index=3)

        with col2:
            st.header("👁️ 图表预览")

            # 绘制图表
            try:
                fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=100)

                x_data = df[x_col]
                marker_val = None if marker == '无' else marker

                # 绘制每条线
                for col in y_cols:
                    ax.plot(x_data, df[col],
                           label=col,
                           linewidth=linewidth,
                           markersize=markersize,
                           linestyle=linestyle,
                           marker=marker_val)

                # 设置标题和标签
                ax.set_title(title, fontsize=title_fontsize, fontweight='bold')
                ax.set_xlabel(xlabel, fontsize=label_fontsize)
                ax.set_ylabel(ylabel, fontsize=label_fontsize)

                # 设置图例
                ax.legend(loc='best', fontsize=legend_fontsize)

                # 设置网格
                if show_grid:
                    ax.grid(True, alpha=grid_alpha)

                # 设置背景色
                fig.patch.set_facecolor(bg_color)
                ax.set_facecolor(bg_color)

                # 设置坐标轴范围
                if not auto_axis:
                    ax.set_xlim(xmin, xmax)
                    ax.set_ylim(ymin, ymax)

                fig.tight_layout()

                # 显示图表
                st.pyplot(fig)

                # 保存按钮
                st.divider()

                col_btn1, col_btn2, col_btn3 = st.columns(3)

                with col_btn1:
                    # PNG下载
                    buf_png = BytesIO()
                    fig.savefig(buf_png, format='png', dpi=dpi, bbox_inches='tight')
                    buf_png.seek(0)
                    st.download_button(
                        label="💾 下载 PNG",
                        data=buf_png,
                        file_name="plot.png",
                        mime="image/png",
                        use_container_width=True
                    )

                with col_btn2:
                    # PDF下载
                    buf_pdf = BytesIO()
                    fig.savefig(buf_pdf, format='pdf', dpi=dpi, bbox_inches='tight')
                    buf_pdf.seek(0)
                    st.download_button(
                        label="📄 下载 PDF",
                        data=buf_pdf,
                        file_name="plot.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )

                with col_btn3:
                    # SVG下载
                    buf_svg = BytesIO()
                    fig.savefig(buf_svg, format='svg', bbox_inches='tight')
                    buf_svg.seek(0)
                    st.download_button(
                        label="🎨 下载 SVG",
                        data=buf_svg,
                        file_name="plot.svg",
                        mime="image/svg+xml",
                        use_container_width=True
                    )

                plt.close(fig)

            except Exception as e:
                st.error(f"❌ 绘图失败: {str(e)}")

    # 页脚
    st.divider()
    with st.expander("💡 使用提示", expanded=False):
        st.markdown("""
        ### 快速开始
        1. **上传文件** - 在左侧上传Excel文件
        2. **选择数据** - 选择X轴和Y轴列
        3. **调整样式** - 自定义标题、字体、颜色等
        4. **下载图片** - 支持PNG、PDF、SVG格式

        ### 推荐设置

        **论文发表**
        - 格式: PDF
        - 标题字号: 16
        - 坐标轴字号: 14
        - 线宽: 2.0
        - 分辨率: 300-600 DPI

        **演示汇报**
        - 格式: PNG
        - 标题字号: 18
        - 线宽: 2.5-3.0
        - 标记点: 6-8
        - 分辨率: 300 DPI
        """)

    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #666;'>"
        "📊 Excel折线图绘制工具 v2.0 | 基于 Streamlit 构建 | "
        "<a href='https://github.com' target='_blank'>查看文档</a>"
        "</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
