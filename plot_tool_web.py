# -*- coding: utf-8 -*-
"""
光谱数据折线图绘制工具 (Web版本)
支持UV-Vis光谱TXT文件批量上传和交互式分析
基于Streamlit的现代化Web界面
"""

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO, StringIO
import os

# 配置matplotlib中文支持
matplotlib.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'SimHei', 'STHeiti', 'DejaVu Sans']
matplotlib.rcParams['axes.unicode_minus'] = False


def parse_txt_file(uploaded_file):
    """
    解析UV-Vis光谱仪TXT文件
    返回DataFrame，包含波长和吸收值两列
    """
    try:
        # 尝试不同的编码读取文件
        encodings = ['gbk', 'utf-8', 'gb2312', 'utf-16']
        content = None

        for encoding in encodings:
            try:
                uploaded_file.seek(0)
                content = uploaded_file.read().decode(encoding)
                break
            except (UnicodeDecodeError, UnicodeError):
                continue

        if content is None:
            st.error(f"无法读取文件 {uploaded_file.name}，编码格式不支持")
            return None

        # 查找数据开始的行（包含"波长"和"吸收"关键词的行）
        lines = content.split('\n')
        data_start_idx = -1

        for idx, line in enumerate(lines):
            if '波长' in line or 'Wavelength' in line.lower():
                data_start_idx = idx + 1
                break

        if data_start_idx == -1:
            st.error(f"文件 {uploaded_file.name} 格式不正确，未找到数据起始行")
            return None

        # 提取数据部分
        data_lines = []
        for line in lines[data_start_idx:]:
            line = line.strip()
            if line and ',' in line:
                parts = line.split(',')
                if len(parts) >= 2:
                    try:
                        wavelength = float(parts[0])
                        absorbance = float(parts[1])
                        data_lines.append([wavelength, absorbance])
                    except ValueError:
                        continue

        if not data_lines:
            st.error(f"文件 {uploaded_file.name} 中未找到有效数据")
            return None

        # 创建DataFrame
        df = pd.DataFrame(data_lines, columns=['Wavelength', 'Absorbance'])
        return df

    except Exception as e:
        st.error(f"解析文件 {uploaded_file.name} 时出错: {str(e)}")
        return None


def merge_txt_files(uploaded_files):
    """
    合并多个TXT文件的数据
    第一列为波长，后续列为各文件的吸收值
    """
    if not uploaded_files:
        return None

    # 解析所有文件
    dfs = []
    file_names = []

    for uploaded_file in uploaded_files:
        df = parse_txt_file(uploaded_file)
        if df is not None:
            # 提取文件名（不含扩展名）作为列名
            file_name = os.path.splitext(uploaded_file.name)[0]
            file_names.append(file_name)
            dfs.append(df)

    if not dfs:
        return None

    # 以第一个文件的波长为基准
    merged_df = dfs[0][['Wavelength']].copy()

    # 合并所有文件的吸收值
    for idx, df in enumerate(dfs):
        # 确保波长匹配（使用merge）
        df_temp = df.copy()
        df_temp.columns = ['Wavelength', file_names[idx]]
        merged_df = pd.merge(merged_df, df_temp, on='Wavelength', how='inner')

    return merged_df


# 页面配置
st.set_page_config(
    page_title="光谱数据绘图工具",
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
    st.markdown('<div class="main-header">📊 光谱数据折线图绘制工具</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">支持UV-Vis光谱TXT文件批量上传和交互式分析</div>', unsafe_allow_html=True)

    # 侧边栏 - 文件上传和数据选择
    with st.sidebar:
        st.header("📁 数据加载")

        uploaded_files = st.file_uploader(
            "选择光谱TXT文件（可多选）",
            type=['txt'],
            accept_multiple_files=True,
            help="支持UV-Vis光谱仪导出的TXT格式文件，可一次上传多个文件"
        )

        if uploaded_files:
            try:
                # 合并所有TXT文件
                df = merge_txt_files(uploaded_files)

                if df is None:
                    st.error("❌ 无法解析文件，请检查文件格式")
                    return

                st.success(f"✅ 成功加载 {len(uploaded_files)} 个文件，共 {len(df)} 个数据点")

                # 显示文件列表
                with st.expander("📄 已加载文件", expanded=False):
                    for idx, file in enumerate(uploaded_files, 1):
                        st.write(f"{idx}. {file.name}")

                # 显示数据预览
                with st.expander("📋 数据预览", expanded=False):
                    st.dataframe(df.head(20))

                st.divider()

                # 列选择
                st.header("🎯 数据选择")

                columns = df.columns.tolist()

                # X轴选择（默认为Wavelength）
                x_col = st.selectbox(
                    "横坐标 (X轴)",
                    columns,
                    index=0,
                    help="通常选择波长作为横坐标"
                )

                # Y轴选择（默认选择所有吸收值列）
                default_y = [col for col in columns if col != x_col]
                y_cols = st.multiselect(
                    "纵坐标 (Y轴，可多选)",
                    columns,
                    default=default_y,
                    help="可以选择多个样品进行对比"
                )

                if not y_cols:
                    st.warning("⚠️ 请至少选择一个Y轴列")
                    return

            except Exception as e:
                st.error(f"❌ 处理文件失败: {str(e)}")
                return
        else:
            st.info("👆 请上传光谱TXT文件开始\n\n支持的文件格式：\n- UV-Vis光谱仪TXT文件\n- 可一次上传多个文件进行对比")
            return

    # 主内容区
    if uploaded_files and y_cols:
        # 创建两列布局
        col1, col2 = st.columns([1, 2])

        with col1:
            st.header("🎨 样式设置")

            # 图表模式选择
            with st.expander("🔧 图表模式", expanded=True):
                chart_mode = st.radio(
                    "选择图表类型",
                    ["交互式图表 (Plotly)", "静态图表 (Matplotlib)"],
                    help="交互式图表支持鼠标框选放大、滚轮缩放、拖动平移等功能，适合数据探索；静态图表适合导出发表"
                )

                if chart_mode == "交互式图表 (Plotly)":
                    st.info("💡 **交互操作说明：**\n"
                           "- 🖱️ **框选放大**：按住鼠标左键拖动选择区域\n"
                           "- 🔍 **滚轮缩放**：鼠标滚轮放大/缩小\n"
                           "- 🔄 **重置视图**：双击图表\n"
                           "- ↔️ **平移**：点击工具栏的平移按钮后拖动")

            # 标题和标签
            with st.expander("📝 标题和标签", expanded=True):
                title = st.text_input("图表标题", value="UV-Vis吸收光谱")
                xlabel = st.text_input("X轴标签", value="Wavelength (nm)")
                ylabel = st.text_input("Y轴标签", value="Absorbance")

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
                if chart_mode == "交互式图表 (Plotly)":
                    # ===== Plotly 交互式图表 =====
                    fig = go.Figure()

                    x_data = df[x_col]

                    # 将 matplotlib 线型转换为 plotly 线型
                    plotly_linestyle = {
                        '-': 'solid',
                        '--': 'dash',
                        '-.': 'dashdot',
                        ':': 'dot'
                    }.get(linestyle, 'solid')

                    # 将 matplotlib 标记转换为 plotly 标记
                    plotly_marker = {
                        '无': None,
                        'o': 'circle',
                        's': 'square',
                        '^': 'triangle-up',
                        'v': 'triangle-down',
                        'D': 'diamond',
                        '*': 'star',
                        '+': 'cross',
                        'x': 'x'
                    }.get(marker, None)

                    # 绘制每条线
                    for col in y_cols:
                        marker_dict = dict(
                            size=markersize * 2,  # Plotly 的标记尺寸约为 Matplotlib 的 2 倍
                            symbol=plotly_marker
                        ) if plotly_marker else None

                        fig.add_trace(go.Scatter(
                            x=x_data,
                            y=df[col],
                            mode='lines+markers' if plotly_marker else 'lines',
                            name=col,
                            line=dict(
                                width=linewidth,
                                dash=plotly_linestyle
                            ),
                            marker=marker_dict
                        ))

                    # 设置布局
                    fig.update_layout(
                        title=dict(
                            text=title,
                            font=dict(size=title_fontsize, family='Arial'),
                            x=0.5,
                            xanchor='center'
                        ),
                        xaxis=dict(
                            title=dict(text=xlabel, font=dict(size=label_fontsize)),
                            showgrid=show_grid,
                            gridcolor=f'rgba(128,128,128,{grid_alpha})',
                            range=[xmin, xmax] if not auto_axis else None
                        ),
                        yaxis=dict(
                            title=dict(text=ylabel, font=dict(size=label_fontsize)),
                            showgrid=show_grid,
                            gridcolor=f'rgba(128,128,128,{grid_alpha})',
                            range=[ymin, ymax] if not auto_axis else None
                        ),
                        legend=dict(
                            font=dict(size=legend_fontsize)
                        ),
                        plot_bgcolor=bg_color,
                        paper_bgcolor=bg_color,
                        height=fig_height * 100,
                        width=fig_width * 100,
                        # 启用交互式缩放功能
                        dragmode='zoom',  # 默认为框选缩放模式
                        hovermode='closest'
                    )

                    # 配置工具栏
                    config = {
                        'displayModeBar': True,
                        'displaylogo': False,
                        'modeBarButtonsToAdd': ['drawopenpath', 'eraseshape'],
                        'modeBarButtonsToRemove': [],
                        'toImageButtonOptions': {
                            'format': 'png',
                            'filename': 'plot',
                            'height': fig_height * dpi,
                            'width': fig_width * dpi,
                            'scale': dpi / 100
                        }
                    }

                    # 显示交互式图表
                    st.plotly_chart(fig, use_container_width=True, config=config)

                    # 下载按钮
                    st.divider()
                    col_btn1, col_btn2 = st.columns(2)

                    with col_btn1:
                        # HTML 下载（包含完整交互功能）
                        html_str = fig.to_html(include_plotlyjs='cdn', config=config)
                        st.download_button(
                            label="💾 下载 HTML (交互式)",
                            data=html_str,
                            file_name="plot_interactive.html",
                            mime="text/html",
                            use_container_width=True,
                            help="下载包含完整交互功能的HTML文件"
                        )

                    with col_btn2:
                        # PNG 下载（使用工具栏的下载按钮）
                        st.info("📸 使用图表右上角工具栏的相机按钮下载PNG图片")

                else:
                    # ===== Matplotlib 静态图表 =====
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
        1. **上传文件** - 在左侧上传UV-Vis光谱TXT文件（可多选）
        2. **自动合并** - 程序自动合并多个文件，波长为X轴，各文件吸收值为Y轴
        3. **选择模式** - 交互式图表（数据探索）或静态图表（论文发表）
        4. **调整样式** - 自定义标题、字体、颜色等
        5. **下载图片** - 支持多种格式

        ### 文件格式说明
        - 支持UV-Vis光谱仪导出的TXT文件
        - 文件中应包含"波长"和"吸收值"数据列
        - 可一次上传多个样品文件进行对比
        - 自动使用文件名作为图例标签

        ### 图表模式选择

        **交互式图表 (Plotly) - 推荐用于光谱数据探索**
        - ✅ 支持鼠标框选局部波段放大
        - ✅ 支持滚轮缩放查看吸收峰细节
        - ✅ 支持拖动平移
        - ✅ 双击重置视图
        - ✅ 悬停显示精确波长和吸收值
        - 📥 下载格式：HTML（保留交互功能）、PNG（通过工具栏）

        **静态图表 (Matplotlib) - 推荐用于论文发表**
        - ✅ 高质量矢量图输出
        - ✅ 精确控制每个细节
        - ✅ 符合期刊要求
        - 📥 下载格式：PNG、PDF、SVG

        ### 推荐设置

        **光谱数据探索**
        - 模式: 交互式图表 (Plotly)
        - 线宽: 2.0
        - 标记点: 无（光谱数据密集）
        - 使用框选放大查看吸收峰细节

        **论文发表**
        - 模式: 静态图表 (Matplotlib)
        - 格式: PDF 或 SVG
        - 标题字号: 16
        - 坐标轴字号: 14
        - 线宽: 1.5-2.0
        - 分辨率: 300-600 DPI

        **多样品对比**
        - 一次上传多个TXT文件
        - 使用不同颜色自动区分
        - 可选择性显示/隐藏某些样品
        - 使用交互式图表便于对比吸收峰位置
        """)

    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #666;'>"
        "📊 光谱数据折线图绘制工具 v4.0 | 支持多TXT文件上传和交互式缩放 | 基于 Streamlit 构建"
        "</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
