import streamlit as st
import pandas as pd
import re
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
import io

# ================= 1. 页面配置 =================
st.set_page_config(
    page_title="财务报告自动化生成器", 
    page_icon="📊",
    layout="wide"
)

st.title("📊 财务分析报告自动化助手")
st.markdown("""
**💡 使用说明：**
1. 上传 **Excel 底稿**（必须）。
2. 上传 **Word 附注**（可选，支持多文件）。
3. 系统会自动计算数据，生成 **数据分析语料**。
4. 表格上方提供 **Word (精排版)** 和 **Excel** 两种格式下载。
""")

# ================= 2. 侧边栏：文件上传 =================
with st.sidebar:
    st.header("📂 请上传文件")
    uploaded_excel = st.file_uploader("1. 上传 Excel 底稿 (必须)", type=["xlsx", "xlsm"])
    uploaded_word_files = st.file_uploader(
        "2. 上传 Word 附注 (可选)", 
        type=["docx"], 
        accept_multiple_files=True,
        help="支持按住 Ctrl/Command 键多选文件，或者多次拖入。"
    )
    st.info("💡 提示：数据只在浏览器本地处理，不会上传给第三方 AI，绝对安全。")
    header_row = st.number_input("Excel表头所在行 (默认2，即第3行)", value=2, min_value=0)

# ================= 3. 核心逻辑函数 =================

def set_cell_border(cell, **kwargs):
    """设置单元格边框"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        if border_name in kwargs:
            edge = kwargs[border_name]
            tcBorders = tcPr.first_child_found_in("w:tcBorders")
            if tcBorders is None:
                tcBorders = OxmlElement('w:tcBorders')
                tcPr.append(tcBorders)
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), edge.get('val', 'single'))
            border.set(qn('w:sz'), str(edge.get('sz', 4)))
            border.set(qn('w:space'), str(edge.get('space', 0)))
            border.set(qn('w:color'), edge.get('color', 'auto'))
            tcBorders.append(border)

def create_word_table_file(df, title="数据表"):
    """🔥 生成精排版 Word 表格"""
    doc = Document()
    
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.font.size = Pt(10.5)

    heading = doc.add_heading(title, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in heading.runs:
        run.font.name = 'SimHei'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        run.font.color.rgb = None

    export_df = df.reset_index()
    table = doc.add_table(rows=1, cols=len(export_df.columns))
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table.autofit = False 
    
    col_widths = [Cm(3.5)] + [Cm(2.2)] * (len(export_df.columns) - 1)
    for i, width in enumerate(col_widths):
        for row in table.rows:
            row.cells[i].width = width

    # --- 表头 ---
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(export_df.columns):
        cell = hdr_cells[i]
        cell.text = str(col_name)
        set_cell_border(cell, top={"val": "single", "sz": 12}, bottom={"val": "single", "sz": 12}, left={"val": "single", "sz": 4}, right={"val": "single", "sz": 4})
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in paragraph.runs:
            run.font.bold = True
            run.font.size = Pt(10.5)
            run.font.name = 'SimHei'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')

    # --- 数据 ---
    for r_idx, row in export_df.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            cell = row_cells[i]
            cell.text = str(val)
            set_cell_border(cell, top={"val": "single", "sz": 4}, bottom={"val": "single", "sz": 4}, left={"val": "single", "sz": 4}, right={"val": "single", "sz": 4})
            if r_idx == len(export_df) - 1:
                 set_cell_border(cell, bottom={"val": "single", "sz": 12})

            paragraph = cell.paragraphs[0]
            if i == 0:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            cell.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER

            for run in paragraph.runs:
                run.font.size = Pt(9)
                run.font.name = 'Times New Roman'
                run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def load_single_word(file_obj):
    try:
        file_obj.seek(0)
        doc = Document(file_obj)
        full_text = [p.text.strip() for p in doc.paragraphs if len(p.text.strip()) > 5]
        return "\n".join(full_text), True 
    except Exception as e:
        if "is not a zip file" in str(e):
            return f"❌ 格式错误：{file_obj.name} 不是标准 .docx，请另存为后上传。", False
        return f"❌ 读取失败：{file_obj.name}", False

def create_excel_file(df):
    """生成 Excel 文件 (保持字符串格式，确保所见即所得)"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='资产明细')
    output.seek(0)
    return output

def find_context(subject, full_text):
    if not full_text: return ""
    clean_sub = subject.replace(" ", "")
    idx = full_text.find(clean_sub)
    if idx == -1: return "（无附注）"
    start = max(0, idx - 600)
    end = min(len(full_text), idx + 1200) 
    return full_text[start:end].replace('\n', ' ')

def clean_date_label(header_str):
    s = str(header_str).replace('\n', '')
    year = re.search(r'(\d{4})', s)
    y_str = year.group(1) if year else "T"
    suffix = "6月末" if ("一期" in s or "6月" in s) else "年末"
    return f"{y_str}年{suffix}"

def safe_pct(num, denom):
    return (num / denom * 100) if denom != 0 else 0.0

# ================= 4. 主程序逻辑 =================

if uploaded_excel:
    
    word_text_all = ""
    if uploaded_word_files:
        for w_file in uploaded_word_files:
            content, success = load_single_word(w_file)
            if success:
                word_text_all += f"\n【来源：{w.name}】\n{content}"
            else:
                st.sidebar.error(content)

    try:
        # 读取数据
        df = pd.read_excel(uploaded_excel, sheet_name='1.合并资产表', header=header_row)
        df = df.iloc[:, [0, 4, 5, 6]]
        orig_cols = df.columns.tolist()
        
        # 动态获取日期列名
        d_t = clean_date_label(orig_cols[1])
        d_t1 = clean_date_label(orig_cols[2])
        d_t2 = clean_date_label(orig_cols[3])
        
        df.columns = ['科目', 'T', 'T_1', 'T_2']
        
        df = df.dropna(subset=['科目'])
        df['科目'] = df['科目'].astype(str).str.strip()
        for c in ['T', 'T_1', 'T_2']:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
        df.set_index('科目', inplace=True)
        
        total_assets = df[df.index.str.contains('资产总计|资产总额')].iloc[0]
        curr_assets = df[df.index.str.contains('流动资产合计')].iloc[0]
        non_curr_assets = df[df.index.str.contains('非流动资产合计')].iloc[0]
        
        # 计算占比
        for period in ['T', 'T_1', 'T_2']:
            total = total_assets[period]
            if total != 0:
                df[f'占比_{period}'] = df[period] / total
            else:
                df[f'占比_{period}'] = 0.0

        # ================= 5. 结果展示 =================
        tab1, tab2, tab3 = st.tabs(["📋 1. 资产明细", "📝 2. 综述文案", "🤖 3. 重点科目分析"])
        
        with tab1:
            c1, c2, c3 = st.columns([6, 1.2, 1.2]) 
            with c1: st.markdown("### 资产结构明细")
            
            # 🔥 核心修改：统一格式化
            display_df = df.copy()
            
            # 1. 格式化金额：带千分位，保留2位小数 (例: 1,234.56)
            for p in ['T', 'T_1', 'T_2']:
                display_df[f'fmt_{p}'] = display_df[p].apply(lambda x: f"{x:,.2f}")
            
            # 2. 格式化占比：乘100，保留2位小数，不带% (例: 12.34)
            for p in ['T', 'T_1', 'T_2']:
                display_df[f'fmt_pct_{p}'] = (display_df[f'占比_{p}'] * 100).apply(lambda x: f"{x:.2f}")

            # 3. 构造最终展示的 DataFrame (重命名+排序)
            final_df = pd.DataFrame(index=display_df.index)
            # T期
            final_df[f"{d_t} 金额"] = display_df['fmt_T']
            final_df[f"{d_t} 占比(%)"] = display_df['fmt_pct_T']
            # T-1期
            final_df[f"{d_t1} 金额"] = display_df['fmt_T_1']
            final_df[f"{d_t1} 占比(%)"] = display_df['fmt_pct_T_1']
            # T-2期
            final_df[f"{d_t2} 金额"] = display_df['fmt_T_2']
            final_df[f"{d_t2} 占比(%)"] = display_df['fmt_pct_T_2']

            with c2:
                doc_file = create_word_table_file(final_df, title="资产结构情况表")
                st.download_button("📥 下载 Word", doc_file, "资产结构明细.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            
            with c3:
                excel_file = create_excel_file(final_df)
                st.download_button("📥 下载 Excel", excel_file, "资产结构明细.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            # 🔥 网页直接展示 final_df，所见即所得
            st.dataframe(final_df, use_container_width=True)

        with tab2:
            st.subheader("资产结构总体分析")
            st.markdown("👇 **直接复制到报告：**")
            top_5 = df.sort_values(by='T', ascending=False).head(5).index.tolist()
            text_overview = (
                f"报告期内，发行人资产总额分别为{total_assets['T_2']:,.2f}万元、{total_assets['T_1']:,.2f}万元和{total_assets['T']:,.2f}万元。\n\n"
                f"其中，流动资产金额分别为{curr_assets['T_2']:,.2f}万元、{curr_assets['T_1']:,.2f}万元和{curr_assets['T']:,.2f}万元，"
                f"占总资产的比例分别为{safe_pct(curr_assets['T_2'], total_assets['T_2']):.2f}%、"
                f"{safe_pct(curr_assets['T_1'], total_assets['T_1']):.2f}%和"
                f"{safe_pct(curr_assets['T'], total_assets['T']):.2f}%；\n\n"
                f"非流动资产金额分别为{non_curr_assets['T_2']:,.2f}万元、{non_curr_assets['T_1']:,.2f}万元和{non_curr_assets['T']:,.2f}万元，"
                f"占总资产的比例分别为{safe_pct(non_curr_assets['T_2'], total_assets['T_2']):.2f}%、"
                f"{safe_pct(non_curr_assets['T_1'], total_assets['T_1']):.2f}%和"
                f"{safe_pct(non_curr_assets['T'], total_assets['T']):.2f}%。\n\n"
                f"在总资产构成中，公司资产主要为 **{'、'.join(top_5)}** 等。"
            )
            st.code(text_overview, language='text')

        with tab3:
            st.subheader("🤖 重点科目分析数据 (Copilot 模式)")
            st.caption("👉 点击代码块右上角的 **📄 复制**，粘贴给 DeepSeek 或 ChatGPT。")
            major_subjects = df[df['占比_T'] > 0.01].index.tolist()
            for subject in major_subjects:
                row = df.loc[subject]
                diff = row['T'] - row['T_1']
                pct = safe_pct(diff, row['T_1'])
                direction = "增加" if diff >= 0 else "减少"
                pct_label = "增幅" if diff >= 0 else "降幅"
                
                prompt_base = f"""【任务】：请分析“{subject}”的变动情况。
【1. 财务具体科目数据 (Trend)】
{d_t2}、{d_t1}及{d_t}，发行人{subject}余额分别为{row['T_2']:,.2f}万元、{row['T_1']:,.2f}万元和{row['T']:,.2f}万元，占总资产的比例分别为{row['占比_T_2']*100:.2f}%、{row['占比_T_1']*100:.2f}%和{row['占比_T']*100:.2f}%。
【2. 财务硬数据变动 (Analysis)】
截至{d_t}，发行人{subject}较{d_t1}{direction}{abs(diff):,.2f}万元，{pct_label}为{abs(pct):.2f}%。"""
                
                context = find_context(subject, word_text_all) if word_text_all else ""
                prompt_final = prompt_base + (f"\n\n【3. Word 附注软信息 (Context)】\n{context}\n\n【4. 写作指令】\n请结合上述数据和附注，分析变动原因（即“主要系...所致”）。如果附注中未提及，请写“主要系业务规模变动所致”。" if context else "")
                
                with st.expander(f"📌 {subject} (占比 {row['占比_T']*100:.2f}%)"):
                    st.code(prompt_final, language='text')

    except Exception as e:
        st.error(f"Excel 解析出错: {e}")
        st.info("请检查 Excel 格式是否与模版一致。")

else:
    st.info("👈 请在左侧上传文件以开始分析")
