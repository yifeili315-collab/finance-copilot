import streamlit as st
import pandas as pd
import re
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
import io

# ================= 1. 页面配置 =================
st.set_page_config(
    page_title="智能财务分析系统", 
    page_icon="📈",
    layout="wide"
)

# ================= 2. 核心逻辑函数 =================

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
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体') 
        run.font.bold = True
        run.font.color.rgb = None

    export_df = df.reset_index()
    table = doc.add_table(rows=1, cols=len(export_df.columns))
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table.autofit = False 
    
    col_widths = [Cm(3.5)] + [Cm(2.2)] * (len(export_df.columns) - 1)
    for i, width in enumerate(col_widths):
        for row in table.rows:
            row.cells[i].width = width

    hdr_cells = table.rows[0].cells
    table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    table.rows[0].height = Cm(0.6)

    for i, col_name in enumerate(export_df.columns):
        cell = hdr_cells[i]
        cell.text = str(col_name)
        top_sz = 12
        bottom_sz = 12 
        left_sz = 12 if i == 0 else 4
        right_sz = 12 if i == len(export_df.columns) - 1 else 4
        set_cell_border(cell, top={"val": "single", "sz": top_sz}, bottom={"val": "single", "sz": bottom_sz}, left={"val": "single", "sz": left_sz}, right={"val": "single", "sz": right_sz})
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.line_spacing = 1.0 
        paragraph.paragraph_format.space_before = Pt(0) 
        paragraph.paragraph_format.space_after = Pt(0)  
        for run in paragraph.runs:
            run.font.bold = True
            run.font.size = Pt(10.5)
            run.font.name = 'Times New Roman'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    for r_idx, row in export_df.iterrows():
        row_cells = table.add_row().cells
        table.rows[r_idx+1].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        table.rows[r_idx+1].height = Cm(0.6)
        subject_name = str(row[0])
        is_bold_row = "合计" in subject_name or "总计" in subject_name
        for i, val in enumerate(row):
            cell = row_cells[i]
            cell.text = str(val)
            bottom_sz = 12 if r_idx == len(export_df) - 1 else 4
            left_sz = 12 if i == 0 else 4
            right_sz = 12 if i == len(export_df.columns) - 1 else 4
            set_cell_border(cell, top={"val": "single", "sz": 4}, bottom={"val": "single", "sz": bottom_sz}, left={"val": "single", "sz": left_sz}, right={"val": "single", "sz": right_sz})
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            paragraph = cell.paragraphs[0]
            if i == 0:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
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
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='数据明细')
    output.seek(0)
    return output

def load_single_word(file_obj):
    try:
        file_obj.seek(0)
        doc = Document(file_obj)
        full_text = [p.text.strip() for p in doc.paragraphs if len(p.text.strip()) > 5]
        return "\n".join(full_text), True, ""
    except Exception as e:
        error_msg = str(e)
        if "is not a zip file" in error_msg:
            friendly_msg = (f"❌ **【格式错误】** 文件：{file_obj.name}\n\n**原因**：这是一个“伪装”的 .docx 文件。\n\n👉 **解决方法：**\n1. 在电脑上用 Word 打开该文件。\n2. 点击左上角【文件】->【另存为】。\n3. 文件类型务必手动选择【Word 文档 (*.docx)】。\n4. 保存后，上传新的文件即可。")
            return "", False, friendly_msg
        else:
            return "", False, f"❌ 读取失败 {file_obj.name}: {error_msg}"

def find_context(subject, word_data_list):
    if not word_data_list: return ""
    clean_sub = subject.replace(" ", "")
    found_contexts = []
    for item in word_data_list:
        content = item['content']
        source = item['source']
        idx = content.find(clean_sub)
        if idx != -1:
            start = max(0, idx - 600)
            end = min(len(content), idx + 1200)
            ctx = content[start:end].replace('\n', ' ')
            found_contexts.append(f"📄 **来源：{source}**\n{ctx}")
    if not found_contexts: return "（未检索到相关附注）"
    return "\n\n".join(found_contexts)

def extract_date_label(header_str):
    s = str(header_str).strip()
    match = re.search(r'[【\[](.*?)[】\]]', s)
    if match: return match.group(1)
    year = re.search(r'(\d{4})', s)
    if year: return f"{year.group(1)}年"
    return s

def safe_pct(num, denom):
    return (num / denom * 100) if denom != 0 else 0.0

# 模糊查找函数
def find_row_fuzzy(df, keywords):
    if isinstance(keywords, str): keywords = [keywords]
    clean_index = df.index.astype(str).str.replace(r'\s+', '', regex=True)
    
    for kw in keywords:
        clean_kw = kw.replace(" ", "")
        mask = clean_index == clean_kw 
        if mask.any():
            return df.loc[df.index[mask][0]]
            
    for kw in keywords:
        clean_kw = kw.replace(" ", "")
        mask = clean_index.str.contains(clean_kw, case=False, na=False)
        if mask.any():
            return df.loc[df.index[mask][0]]

    raise ValueError(f"未找到包含 {' / '.join(keywords)} 的行")

def process_analysis_tab(df_raw, word_data_list, total_col_name, analysis_name, d_labels):
    try:
        # 🔥 核心修正：负债结构分析的精准切片
        if analysis_name == "负债":
             # 1. 使用正则表达式精准定位“负债合计”
             # ^ 表示开始, $ 表示结束, \s* 表示允许有空格
             # 这样就能排除 "流动负债合计" (前面有字)
             # 我们在 index 中搜索匹配这个模式的行
             
             # 先把 index 转成 string
             index_series = df_raw.index.astype(str)
             
             # 查找完全匹配 "负债合计" (忽略前后空格) 的行
             # 如果你的表里写的是 "负 债 合 计"，我们需要先去除空格再匹配，或者用宽容正则
             
             # 方案：先创建一个没有空格的 index 映射
             clean_index = index_series.str.replace(r'\s+', '', regex=True)
             clean_target = total_col_name.replace(" ", "") # "负债合计"
             
             match_mask = (clean_index == clean_target)
             
             if match_mask.any():
                 # 获取匹配行的 Label
                 target_label = df_raw.index[match_mask][0]
                 
                 # 获取行号
                 idx_pos = df_raw.index.get_loc(target_label)
                 
                 # 如果有重复(比如母公司/合并)，通常取最后一个（或者看需求）
                 # 这里假设我们已经读了合并表，取最后一个通常比较安全（因为总计在最下）
                 if isinstance(idx_pos, slice):
                     idx_pos = idx_pos.stop - 1
                 elif hasattr(idx_pos, '__iter__'): 
                     idx_pos = idx_pos[-1]
                 
                 # 🔥 执行切片：只保留到“负债合计”这一行
                 if isinstance(idx_pos, int):
                    df_raw = df_raw.iloc[:idx_pos + 1]
             else:
                 st.warning(f"⚠️ 未找到严格等于 '{total_col_name}' 的行，将显示完整表格。建议检查 Excel 行名。")

        # 2. 获取总计数据
        total_row = find_row_fuzzy(df_raw, [total_col_name])
        
    except Exception as e:
        st.error(f"❌ 数据处理错误: {e}")
        return

    # 3. 计算占比
    df = df_raw.copy()
    for period in ['T', 'T_1', 'T_2']:
        total = total_row[period]
        if total != 0:
            df[f'占比_{period}'] = df[period] / total
        else:
            df[f'占比_{period}'] = 0.0

    tab1, tab2, tab3 = st.tabs(["📋 明细数据", "📝 综述文案", "🤖 AI 分析指令"])

    # 4. 显示明细数据 (现在是切片后的干净表格了！)
    with tab1:
        c1, c2, c3 = st.columns([6, 1.2, 1.2]) 
        with c1: st.markdown(f"### {analysis_name}结构明细")
        
        display_df = df.copy()
        for p in ['T', 'T_1', 'T_2']:
            display_df[f'fmt_{p}'] = display_df[p].apply(lambda x: f"{x:,.2f}")
            display_df[f'fmt_pct_{p}'] = (display_df[f'占比_{p}'] * 100).apply(lambda x: f"{x:.2f}")

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

    with tab2:
        st.markdown("👇 **直接复制：**")
        top_5 = df.sort_values(by='T', ascending=False).head(5).index.tolist()
        
        text = ""
        try:
            if analysis_name == "资产":
                curr_row = find_row_fuzzy(df_raw, ['流动资产合计', '流动资产小计'])
                non_curr_row = find_row_fuzzy(df_raw, ['非流动资产合计', '非流动资产小计'])
                
                text = (f"报告期内，发行人资产总额分别为{total_row['T_2']:,.2f}万元、{total_row['T_1']:,.2f}万元和{total_row['T']:,.2f}万元。\n\n"
                        f"其中，流动资产金额分别为{curr_row['T_2']:,.2f}万元、{curr_row['T_1']:,.2f}万元和{curr_row['T']:,.2f}万元，"
                        f"占总资产的比例分别为{safe_pct(curr_row['T_2'], total_row['T_2']):.2f}%、{safe_pct(curr_row['T_1'], total_row['T_1']):.2f}%和{safe_pct(curr_row['T'], total_row['T']):.2f}%；\n\n"
                        f"非流动资产金额分别为{non_curr_row['T_2']:,.2f}万元、{non_curr_row['T_1']:,.2f}万元和{non_curr_row['T']:,.2f}万元，"
                        f"占总资产的比例分别为{safe_pct(non_curr_row['T_2'], total_row['T_2']):.2f}%、{safe_pct(non_curr_row['T_1'], total_row['T_1']):.2f}%和{safe_pct(non_curr_row['T'], total_row['T']):.2f}%。\n\n"
                        f"在总资产构成中，公司资产主要为 **{'、'.join(top_5)}** 等。")
            elif analysis_name == "负债":
                curr_row = find_row_fuzzy(df_raw, ['流动负债合计', '流动负债小计'])
                non_curr_row = find_row_fuzzy(df_raw, ['非流动负债合计', '非流动负债小计'])
                
                diff_prev = total_row['T_1'] - total_row['T_2']
                pct_prev = safe_pct(diff_prev, total_row['T_2'])
                dir_prev = "增加" if diff_prev >= 0 else "减少"
                label_prev = "增幅" if diff_prev >= 0 else "降幅"
                
                diff_curr = total_row['T'] - total_row['T_1']
                pct_curr = safe_pct(diff_curr, total_row['T_1'])
                dir_curr = "增加" if diff_curr >= 0 else "减少"
                label_curr = "增幅" if diff_curr >= 0 else "降幅"
                
                trend_desc = "增长" if diff_curr >= 0 else "下降"

                text = (
                    f"报告期内，发行人负债总额分别为{total_row['T_2']:,.2f}万元、{total_row['T_1']:,.2f}万元和{total_row['T']:,.2f}万元，"
                    f"{d_labels[1]}较{d_labels[2]}{dir_prev}{abs(diff_prev):,.2f}万元，{label_prev}{abs(pct_prev):.2f}%，"
                    f"{d_labels[0]}发行人负债较{d_labels[1]}{dir_curr}{abs(diff_curr):,.2f}万元，{label_curr}{abs(pct_curr):.2f}%。"
                    f"报告期内发行人的负债规模呈现{trend_desc}态势，主要原因为发行人（用户自行分析）。\n\n"
                    
                    f"从负债结构来看，报告期内，流动负债分别为{curr_row['T_2']:,.2f}万元、{curr_row['T_1']:,.2f}万元和{curr_row['T']:,.2f}万元，"
                    f"占负债总额比例分别为{safe_pct(curr_row['T_2'], total_row['T_2']):.2f}%、"
                    f"{safe_pct(curr_row['T_1'], total_row['T_1']):.2f}%和"
                    f"{safe_pct(curr_row['T'], total_row['T']):.2f}%，"
                    f"主要由 **{'、'.join(top_5)}** 等构成；\n\n"
                    
                    f"非流动负债分别为{non_curr_row['T_2']:,.2f}万元、{non_curr_row['T_1']:,.2f}万元和{non_curr_row['T']:,.2f}万元，"
                    f"占负债总额比例分别为{safe_pct(non_curr_row['T_2'], total_row['T_2']):.2f}%、"
                    f"{safe_pct(non_curr_row['T_1'], total_row['T_1']):.2f}%和"
                    f"{safe_pct(non_curr_row['T'], total_row['T']):.2f}%。"
                )
            else:
                text = f"报告期内，发行人{analysis_name}总额分别为{total_row['T_2']:,.2f}万元、{total_row['T_1']:,.2f}万元和{total_row['T']:,.2f}万元。\n主要构成项目包括：**{'、'.join(top_5)}** 等。"
        except Exception as e:
             text = f"⚠️ 生成文案时出错: {e}。\n\n请检查您的 Excel 表格中是否包含 **【流动负债合计】** 和 **【非流动负债合计】** 这两行。"
        st.code(text, language='text')

    with tab3:
        if word_data_list:
            st.info(f"💡 **提示**：已结合 Excel 数据与 **{len(word_data_list)} 个 Word 附注** 生成深度分析指令。")
        else:
            st.info(f"💡 **提示**：仅基于 Excel 数据生成指令（未检测到 Word 附注，已自动隐藏“附注线索”部分）。")
            
        st.caption("👉 点击右上角复制，发送给 AI (DeepSeek/ChatGPT)。")
        exclude_list = ['合计', '总计', '总额']
        major_subjects = df[(df['占比_T'] > 0.01) & (~df.index.str.contains('|'.join(exclude_list)))].index.tolist()
        
        denom_text = "总资产" if analysis_name == "资产" else f"{analysis_name}总额"

        for subject in major_subjects:
            row = df.loc[subject]
            
            diff_prev = row['T_1'] - row['T_2']
            pct_prev = safe_pct(diff_prev, row['T_2'])
            dir_prev = "增加" if diff_prev >= 0 else "减少"
            label_prev = "增幅" if diff_prev >= 0 else "降幅"

            diff_curr = row['T'] - row['T_1']
            pct_curr = safe_pct(diff_curr, row['T_1'])
            dir_curr = "增加" if diff_curr >= 0 else "减少"
            label_curr = "增幅" if diff_curr >= 0 else "降幅"
            
            prompt = f"""【任务】分析“{subject}”变动原因。
【1. 数据趋势】
{d_t2}、{d_t1}及{d_t}，发行人{subject}余额分别为{row['T_2']:,.2f}万元、{row['T_1']:,.2f}万元和{row['T']:,.2f}万元，占{denom_text}的比例分别为{row['占比_T_2']*100:.2f}%、{row['占比_T_1']*100:.2f}%和{row['占比_T']*100:.2f}%。
【2. 变动情况】
截至{d_t1}，发行人{subject}较{d_t2}{dir_prev}{abs(diff_prev):,.2f}万元，{label_prev}{abs(pct_prev):.2f}%。
截至{d_t}，发行人{subject}较{d_t1}{dir_curr}{abs(diff_curr):,.2f}万元，{label_curr}{abs(pct_curr):.2f}%。"""

            if word_data_list:
                prompt += f"""
【3. 附注线索】
{find_context(subject, word_data_list)}
【4. 写作要求】
结合数据和附注分析原因。如附注未提及，写“主要系业务规模变动所致”。"""
            
            with st.expander(f"📌 {subject} (占比 {row['占比_T']:.2%})"):
                st.code(prompt, language='text')

# ================= 3. 侧边栏 =================
with st.sidebar:
    st.title("🎛️ 操控台")
    analysis_page = st.radio("请选择要生成的章节：", ["(一) 资产结构分析", "(二) 负债结构分析", "(三) 现金流量分析 (开发中...)", "(四) 财务指标分析 (开发中...)"])
    st.markdown("---")
    
    uploaded_excel = st.file_uploader("Excel 底稿 (必须)", type=["xlsx", "xlsm"])
    uploaded_word_files = st.file_uploader("Word 附注 (可选)", type=["docx"], accept_multiple_files=True)
    
    with st.expander("⚙️ 高级设置 (Sheet名称/表头行)"):
        header_row = st.number_input("表头所在行 (默认2，即第3行)", value=2, min_value=0)
        sheet_asset = st.text_input("资产表 Sheet 名", value="1.合并资产表")
        sheet_liab = st.text_input("负债表 Sheet 名", value="2.合并负债及权益表")

# ================= 4. 主程序 =================

if not uploaded_excel:
    st.title("📊 财务分析报告自动化助手")
    st.info("💡 本系统专为 **公司标准审计底稿模版** 设计，请勿随意修改 Excel 格式。")
    
    st.markdown("""
    ### 🛑 使用前必读 (Requirements)
    为了确保数据读取准确，您的 Excel 文件 **必须** 满足以下条件：
    
    1.  **Sheet 名称严格匹配**：
        * 资产表 -> `1.合并资产表`
        * 负债表 -> `2.合并负债及权益表`
    2.  **数据列位置固定**：系统默认读取 **E、F、G 列**（模版中的“万元”列）。
    3.  **表头位置固定**：表头必须位于 **第 3 行**（即 Excel 左侧行号为 3）。
    
    > **💡 小技巧：如何自定义日期名称？**
    > 系统会自动提取 Excel 表头中 **【 】** 里的文字。
    > * 如果您希望文案显示 **“2023年末”**，请直接将 Excel 表头改为 `【2023年末】`。
    > * 如果您希望文案显示 **“2025年9月末”**，请将 Excel 表头改为 `【2025年9月末】`。
    
    ---
    ### 🚀 快速上手：
    1.  **左侧上传**：拖入 Excel 底稿和 Word 附注。
    2.  **自动分析**：上传即算，点击上方标签页切换 **数据表 / 文案 / AI指令**。
    3.  **一键导出**：支持导出 **精排版 Word 表格** (宋体/加粗/1.5磅边框)。
    """)
    
    st.warning("👈 请先在左侧侧边栏上传 Excel 文件以开始使用。")

else:
    word_data_list = []
    word_error_msgs = []
    if uploaded_word_files:
        for w in uploaded_word_files:
            content, success, err_msg = load_single_word(w) 
            if success:
                word_data_list.append({'source': w.name, 'content': content})
            else:
                word_error_msgs.append(err_msg)
    if word_error_msgs:
        for msg in word_error_msgs: st.error(msg)
    elif uploaded_word_files: st.success(f"✅ 成功读取 {len(word_data_list)} 个 Word 文件！")

    # 🔥 核心修正：模糊查找 Sheet 名称
    def fuzzy_load_excel(file_obj, sheet_name, header_row):
        xl = pd.ExcelFile(file_obj)
        all_sheet_names = xl.sheet_names
        
        if sheet_name in all_sheet_names:
            return pd.read_excel(file_obj, sheet_name=sheet_name, header=header_row), None
        
        clean_target = sheet_name.replace(" ", "")
        for actual_name in all_sheet_names:
            if actual_name.replace(" ", "") == clean_target:
                st.toast(f"⚠️ 检测到 Sheet 名称不一致，已自动修正为：'{actual_name}'")
                return pd.read_excel(file_obj, sheet_name=actual_name, header=header_row), None
        
        return None, all_sheet_names

    def get_clean_data(target_sheet_name):
        try:
            df, all_sheets_if_failed = fuzzy_load_excel(uploaded_excel, target_sheet_name, header_row)
            
            if df is None:
                return None, None, f"未找到 Sheet '{target_sheet_name}' (现有 Sheet: {all_sheets_if_failed})"

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

    st.header(f"📊 {analysis_page}")

    if analysis_page == "(一) 资产结构分析":
        df_asset, d_labels, err = get_clean_data(sheet_asset)
        if df_asset is not None:
            process_analysis_tab(df_asset, word_data_list, "资产总计", "资产", d_labels)
        else:
            st.error(f"❌ 读取失败：{err}\n\n请检查侧边栏【高级设置】中的 Sheet 名称。")

    elif analysis_page == "(二) 负债结构分析":
        df_liab, d_labels, err = get_clean_data(sheet_liab)
        if df_liab is not None:
            total_name = "负债合计" 
            if not df_liab.index.str.contains(total_name).any():
                total_name = "负债总计"
            process_analysis_tab(df_liab, word_data_list, total_name, "负债", d_labels)
        else:
            st.error(f"❌ 读取失败：{err}\n\n请检查侧边栏【高级设置】中的 Sheet 名称。")

    else:
        st.info("🚧 该模块正在施工中，敬请期待后续更新...")
