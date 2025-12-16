import streamlit as st
import pandas as pd
import re
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
import io

# ================= 1. 页面配置 =================
st.set_page_config(
    page_title="智能财务分析系统", 
    page_icon="📈",
    layout="wide"
)

# ================= 2. 核心逻辑函数 (通用工具箱) =================

def set_cell_border(cell, **kwargs):
    """Word表格边框设置"""
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

    # 表头
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(export_df.columns):
        cell = hdr_cells[i]
        cell.text = str(col_name)
        set_cell_border(cell, top={"val": "single", "sz": 12}, bottom={"val": "single", "sz": 12}, left={"val": "single", "sz": 4}, right={"val": "single", "sz": 4})
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in paragraph.runs:
            run.font.bold = True
            run.font.size = Pt(10.5)
            run.font.name = 'SimHei'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')

    # 数据
    for r_idx, row in export_df.iterrows():
        row_cells = table.add_row().cells
        
        subject_name = str(row[0])
        is_bold_row = "合计" in subject_name or "总计" in subject_name

        for i, val in enumerate(row):
            cell = row_cells[i]
            cell.text = str(val)
            set_cell_border(cell, top={"val": "single", "sz": 4}, bottom={"val": "single", "sz": 4}, left={"val": "single", "sz": 4}, right={"val": "single", "sz": 4})
            if r_idx == len(export_df) - 1:
                 set_cell_border(cell, bottom={"val": "single", "sz": 12})
            
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            paragraph = cell.paragraphs[0]
            if i == 0:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            paragraph.paragraph_format.space_before = Pt(2)
            paragraph.paragraph_format.space_after = Pt(2)

            for run in paragraph.runs:
                run.font.size = Pt(9)
                run.font.name = 'Times New Roman'
                run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                if is_bold_row:
                    run.font.bold = True
            
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def create_excel_file(df):
    """生成 Excel"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='数据明细')
    output.seek(0)
    return output

def load_single_word(file_obj):
    """读取 Word (含格式错误拦截)"""
    try:
        file_obj.seek(0)
        doc = Document(file_obj)
        full_text = [p.text.strip() for p in doc.paragraphs if len(p.text.strip()) > 5]
        return "\n".join(full_text), True, ""
    except Exception as e:
        error_msg = str(e)
        if "is not a zip file" in error_msg:
            # 返回友好的错误提示
            friendly_msg = (
                f"❌ 【格式错误】文件：{file_obj.name}\n"
                f"原因：这是一个“伪装”的 .docx 文件（本质可能是老版本 .doc 或其他格式）。\n"
                f"👉 解决方法：\n"
                f"1. 在电脑上用 Word 打开该文件。\n"
                f"2. 点击左上角【文件】->【另存为】。\n"
                f"3. 文件类型务必手动选择【Word 文档 (*.docx)】。\n"
                f"4. 保存后，上传新的文件即可。"
            )
            return "", False, friendly_msg
        else:
            return "", False, f"❌ 读取失败 {file_obj.name}: {error_msg}"

def find_context(subject, full_text):
    """RAG 检索"""
    if not full_text: return ""
    clean_sub = subject.replace(" ", "")
    idx = full_text.find(clean_sub)
    if idx == -1: return "（未检索到相关附注）"
    start = max(0, idx - 600)
    end = min(len(full_text), idx + 1200) 
    return full_text[start:end].replace('\n', ' ')

def extract_date_label(header_str):
    """智能提取日期标签"""
    s = str(header_str).strip()
    match = re.search(r'[【\[](.*?)[】\]]', s)
    if match: return match.group(1)
    year = re.search(r'(\d{4})', s)
    if year: return f"{year.group(1)}年"
    return s

def safe_pct(num, denom):
    return (num / denom * 100) if denom != 0 else 0.0

def process_analysis_tab(df_raw, word_text, total_col_name, analysis_name, d_labels):
    """通用分析函数"""
    # 提取关键行
    try:
        total_row = df_raw[df_raw.index.str.contains(total_col_name)].iloc[0]
    except:
        st.error(f"❌ 分析中断：在表中未找到 '{total_col_name}' 行，请检查 Excel 科目名称或 Sheet 选择是否正确。")
        return

    # 计算占比
    df = df_raw.copy()
    for period in ['T', 'T_1', 'T_2']:
        total = total_row[period]
        if total != 0:
            df[f'占比_{period}'] = df[period] / total
        else:
            df[f'占比_{period}'] = 0.0

    # === 展示界面 ===
    tab1, tab2, tab3 = st.tabs(["📋 明细数据", "📝 综述文案", "🤖 AI 分析指令"])

    # 1. 明细表
    with tab1:
        c1, c2, c3 = st.columns([6, 1.2, 1.2]) 
        with c1: st.markdown(f"### {analysis_name}结构明细")
        
        # 格式化数据
        display_df = df.copy()
        for p in ['T', 'T_1', 'T_2']:
            display_df[f'fmt_{p}'] = display_df[p].apply(lambda x: f"{x:,.2f}")
            display_df[f'fmt_pct_{p}'] = (display_df[f'占比_{p}'] * 100).apply(lambda x: f"{x:.2f}")

        # 构造最终表格
        d_t, d_t1, d_t2 = d_labels
        final_df = pd.DataFrame(index=display_df.index)
        final_df[f"{d_t}"] = display_df['fmt_T']
        final_df["占比(%) "] = display_df['fmt_pct_T']
        final_df[f"{d_t1}"] = display_df['fmt_T_1']
        final_df["占比(%)"] = display_df['fmt_pct_T_1']
        final_df[f"{d_t2}"] = display_df['fmt_T_2']
        final_df[" 占比(%)"] = display_df['fmt_pct_T_2']

        with c2:
            doc_file = create_word_table_file(final_df, title=f"{analysis_name}结构情况表")
            st.download_button(f"📥 下载 Word", doc_file, f"{analysis_name}明细.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        with c3:
            excel_file = create_excel_file(final_df)
            st.download_button(f"📥 下载 Excel", excel_file, f"{analysis_name}明细.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        st.dataframe(final_df, use_container_width=True)

    # 2. 综述文案
    with tab2:
        st.markdown("👇 **直接复制：**")
        top_5 = df.sort_values(by='T', ascending=False).head(5).index.tolist()
        text = ""
        try:
            if analysis_name == "资产":
                curr_row = df_raw[df_raw.index.str.contains('流动资产合计')].iloc[0]
                non_curr_row = df_raw[df_raw.index.str.contains('非流动资产合计')].iloc[0]
                
                text = (
                    f"报告期内，发行人资产总额分别为{total_row['T_2']:,.2f}万元、{total_row['T_1']:,.2f}万元和{total_row['T']:,.2f}万元。\n\n"
                    f"其中，流动资产金额分别为{curr_row['T_2']:,.2f}万元、{curr_row['T_1']:,.2f}万元和{curr_row['T']:,.2f}万元，"
                    f"占总资产的比例分别为{safe_pct(curr_row['T_2'], total_row['T_2']):.2f}%、"
                    f"{safe_pct(curr_row['T_1'], total_row['T_1']):.2f}%和"
                    f"{safe_pct(curr_row['T'], total_row['T']):.2f}%；\n\n"
                    f"非流动资产金额分别为{non_curr_row['T_2']:,.2f}万元、{non_curr_row['T_1']:,.2f}万元和{non_curr_row['T']:,.2f}万元，"
                    f"占总资产的比例分别为{safe_pct(non_curr_row['T_2'], total_row['T_2']):.2f}%、"
                    f"{safe_pct(non_curr_row['T_1'], total_row['T_1']):.2f}%和"
                    f"{safe_pct(non_curr_row['T'], total_row['T']):.2f}%。\n\n"
                    f"在总资产构成中，公司资产主要为 **{'、'.join(top_5)}** 等。"
                )
            elif analysis_name == "负债":
                curr_row = df_raw[df_raw.index.str.contains('流动负债合计')].iloc[0]
                non_curr_row = df_raw[df_raw.index.str.contains('非流动负债合计')].iloc[0]
                
                text = (
                    f"报告期内，发行人负债总额分别为{total_row['T_2']:,.2f}万元、{total_row['T_1']:,.2f}万元和{total_row['T']:,.2f}万元。\n\n"
                    f"其中，流动负债金额分别为{curr_row['T_2']:,.2f}万元、{curr_row['T_1']:,.2f}万元和{curr_row['T']:,.2f}万元，"
                    f"占负债总额的比例分别为{safe_pct(curr_row['T_2'], total_row['T_2']):.2f}%、"
                    f"{safe_pct(curr_row['T_1'], total_row['T_1']):.2f}%和"
                    f"{safe_pct(curr_row['T'], total_row['T']):.2f}%；\n\n"
                    f"非流动负债金额分别为{non_curr_row['T_2']:,.2f}万元、{non_curr_row['T_1']:,.2f}万元和{non_curr_row['T']:,.2f}万元，"
                    f"占负债总额的比例分别为{safe_pct(non_curr_row['T_2'], total_row['T_2']):.2f}%、"
                    f"{safe_pct(non_curr_row['T_1'], total_row['T_1']):.2f}%和"
                    f"{safe_pct(non_curr_row['T'], total_row['T']):.2f}%。\n\n"
                    f"从结构来看，主要构成项目包括：**{'、'.join(top_5)}** 等。"
                )
            else:
                text = f"报告期内，发行人{analysis_name}总额分别为{total_row['T_2']:,.2f}万元、{total_row['T_1']:,.2f}万元和{total_row['T']:,.2f}万元。\n主要构成项目包括：**{'、'.join(top_5)}** 等。"
        except:
             text = f"报告期内，发行人{analysis_name}总额分别为{total_row['T_2']:,.2f}万元、{total_row['T_1']:,.2f}万元和{total_row['T']:,.2f}万元。\n主要构成项目包括：**{'、'.join(top_5)}** 等。"
        
        st.code(text, language='text')

    # 3. AI 指令
    with tab3:
        st.caption("👉 点击右上角复制，发送给 AI (DeepSeek/ChatGPT)。")
        
        exclude_list = ['合计', '总计', '总额']
        major_subjects = df[
            (df['占比_T'] > 0.01) & 
            (~df.index.str.contains('|'.join(exclude_list)))
        ].index.tolist()
        
        for subject in major_subjects:
            row = df.loc[subject]
            diff = row['T'] - row['T_1']
            pct = safe_pct(diff, row['T_1'])
            direction = "增加" if diff >= 0 else "减少"
            pct_label = "增幅" if diff >= 0 else "降幅"
            
            prompt = f"""【任务】分析“{subject}”变动原因。
【1. 数据趋势】
{d_t2}、{d_t1}及{d_t}，余额分别为{row['T_2']:,.2f}万元、{row['T_1']:,.2f}万元和{row['T']:,.2f}万元，占比分别为{row['占比_T_2']*100:.2f}%、{row['占比_T_1']*100:.2f}%和{row['占比_T']*100:.2f}%。
【2. 变动情况】
截至{d_t}，较上期{direction}{abs(diff):,.2f}万元，{pct_label}{abs(pct):.2f}%。
【3. 附注线索】
{find_context(subject, word_text)}
【4. 写作要求】
结合数据和附注分析原因。如附注未提及，写“主要系业务规模变动所致”。"""
            
            with st.expander(f"📌 {subject} (占比 {row['占比_T']:.2%})"):
                st.code(prompt, language='text')


# ================= 3. Sidebar: Global Control =================
with st.sidebar:
    st.title("🎛️ 操控台")
    
    analysis_page = st.radio(
        "请选择要生成的章节：",
        ["(一) 资产结构分析", "(二) 负债结构分析", "(三) 现金流量分析 (开发中...)", "(四) 财务指标分析 (开发中...)"]
    )
    st.markdown("---")
    
    uploaded_excel = st.file_uploader("Excel 底稿 (必须)", type=["xlsx", "xlsm"])
    uploaded_word_files = st.file_uploader("Word 附注 (可选)", type=["docx"], accept_multiple_files=True)
    
    header_row = st.number_input("表头所在行 (默认2)", value=2)
    st.markdown("### 3. Excel Sheet 匹配")
    sheet_asset = st.text_input("资产表 Sheet 名", value="1.合并资产表")
    sheet_liab = st.text_input("负债表 Sheet 名", value="2.合并负债表") 

# ================= 4. Main Logic =================

# 🔥 核心状态切换逻辑
if not uploaded_excel:
    # 状态 A：未上传文件 -> 显示首页引导 (参考了你的截图)
    st.title("📊 财务分析报告自动化助手")
    
    st.markdown("""
    ### 💡 使用说明：
    1. **上传 Excel 底稿 (必须)**：请在左侧侧边栏上传。
    2. **上传 Word 附注 (可选)**：支持上传多个 Word 文件，用于生成原因分析。
    3. **自动计算与生成**：系统会自动提取数据，生成 **数据表格**、**综述文案** 和 **AI 指令**。
    4. **一键导出**：支持导出 **精排版 Word 表格**，直接粘贴到报告中。
    """)
    
    st.info("👈 请先在左侧侧边栏上传 Excel 文件以开始使用。")

else:
    # 状态 B：已上传文件 -> 显示分析结果
    
    # 1. 预处理 Word (含报错显示)
    word_text_all = ""
    word_error_msgs = []
    
    if uploaded_word_files:
        for w in uploaded_word_files:
            content, success, err_msg = load_single_word(w) 
            if success:
                word_text_all += f"\n【来源：{w.name}】\n{content}"
            else:
                word_error_msgs.append(err_msg)
    
    # 🔥 如果有 Word 错误，在主界面顶部醒目显示
    if word_error_msgs:
        for msg in word_error_msgs:
            st.error(msg)
    elif uploaded_word_files:
        st.success("✅ 所有 Excel 和 Word 文件均读取成功！")

    # 2. 通用 Excel 读取器 (含报错显示)
    def get_clean_data(sheet_name):
        try:
            df = pd.read_excel(uploaded_excel, sheet_name=sheet_name, header=header_row)
            df = df.iloc[:, [0, 4, 5, 6]]
            orig_cols = df.columns.tolist()
            d_labels = [extract_date_label(orig_cols[1]), extract_date_label(orig_cols[2]), extract_date_label(orig_cols[3])]
            
            df.columns = ['科目', 'T', 'T_1', 'T_2']
            df = df.dropna(subset=['科目'])
            df['科目'] = df['科目'].astype(str).str.strip()
            for c in ['T', 'T_1', 'T_2']:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
            df.set_index('科目', inplace=True)
            return df, d_labels, None
        except Exception as e:
            return None, None, str(e)

    # 3. 页面路由逻辑
    st.header(f"📊 {analysis_page}")

    if analysis_page == "(一) 资产结构分析":
        df_asset, d_labels, err = get_clean_data(sheet_asset)
        if df_asset is not None:
            process_analysis_tab(df_asset, word_text_all, "资产总计", "资产", d_labels)
        else:
            st.error(f"❌ 读取 Excel 失败：{err}\n请检查【资产表 Sheet 名】是否为：{sheet_asset}")

    elif analysis_page == "(二) 负债结构分析":
        df_liab, d_labels, err = get_clean_data(sheet_liab)
        if df_liab is not None:
            total_name = "负债合计" 
            if not df_liab.index.str.contains(total_name).any():
                total_name = "负债总计"
            process_analysis_tab(df_liab, word_text_all, total_name, "负债", d_labels)
        else:
            st.error(f"❌ 读取 Excel 失败：{err}\n请检查【负债表 Sheet 名】是否为：{sheet_liab}")

    else:
        st.info("🚧 该模块正在施工中，敬请期待后续更新...")
