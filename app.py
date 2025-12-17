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
    """🔥 生成精排版 Word 表格 (宋体+粗边框)"""
    doc = Document()
    
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.font.size = Pt(10.5)

    heading = doc.add_heading(title, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in heading.runs:
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体') # 标题用宋体加粗
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

    # --- 表头 (宋体加粗) ---
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

        set_cell_border(cell, 
                        top={"val": "single", "sz": top_sz}, 
                        bottom={"val": "single", "sz": bottom_sz}, 
                        left={"val": "single", "sz": left_sz}, 
                        right={"val": "single", "sz": right_sz})
        
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

    # --- 数据填充 ---
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
            
            set_cell_border(cell, 
                            top={"val": "single", "sz": 4}, 
                            bottom={"val": "single", "sz": bottom_sz}, 
                            left={"val": "single", "sz": left_sz}, 
                            right={"val": "single", "sz": right_sz})
            
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
    """读取 Word 返回内容"""
    try:
        file_obj.seek(0)
        doc = Document(file_obj)
        full_text = [p.text.strip() for p in doc.paragraphs if len(p.text.strip()) > 5]
        return "\n".join(full_text), True, ""
    except Exception as e:
        error_msg = str(e)
        if "is not a zip file" in error_msg:
            friendly_msg = (
                f"❌ **【格式错误】** 文件：{file_obj.name}\n\n"
                f"**原因**：这是一个“伪装”的 .docx 文件。\n\n"
                f"👉 **解决方法：**\n"
                f"1. 在电脑上用 Word 打开该文件。\n"
                f"2. 点击左上角【文件】->【另存为】。\n"
                f"3. 文件类型务必手动选择【Word 文档 (*.docx)】。\n"
                f"4. 保存后，上传新的文件即可。"
            )
            return "", False, friendly_msg
        else:
            return "", False, f"❌ 读取失败 {file_obj.name}: {error_msg}"

def find_context(subject, word_data_list):
    """
    🔥 多文件 RAG 检索
    word_data_list: [{'source': '文件名', 'content': '内容'}, ...]
    """
    if not word_data_list: return ""
    
    clean_sub = subject.replace(" ", "")
    found_contexts = []
    
    for item in word_data_list:
        content = item['content']
        source = item['source']
        
        idx = content.find(clean_sub)
        if idx != -1:
            # 找到关键词，截取前后文
            start = max(0, idx - 600)
            end = min(len(content), idx + 1200)
            ctx = content[start:end].replace('\n', ' ')
            # 🔥 加上来源标记
            found_contexts.append(f"📄 **来源：{source}**\n{ctx}")
            
    if not found_contexts:
        return "（未检索到相关附注）"
    
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

def process_analysis_tab(df_raw, word_data_list, total_col_name, analysis_name, d_labels):
    """核心分析函数"""
    try:
        total_row = df_raw[df_raw.index.str.contains(total_col_name)].iloc[0]
    except:
        st.error(f"❌ 分析中断：在表中未找到 '{total_col_name}' 行，请检查 Excel 科目名称或 Sheet 选择是否正确。")
        return

    df = df_raw.copy()
    for period in ['T', 'T_1', 'T_2']:
        total = total_row[period]
        if total != 0:
            df[f'占比_{period}'] = df[period] / total
        else:
            df[f'占比_{period}'] = 0.0

    tab1, tab2, tab3 = st.tabs(["📋 明细数据", "📝 综述文案", "🤖 AI 分析指令"])

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

    with tab3:
        st.info(f"💡 **提示**：以下是基于 **{d_t} (最新一期)** 占比前列的科目生成的分析指令。")
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
    
    # 🔥 新增：数据列模式选择
    data_col_mode = st.radio(
        "📊 数据列读取模式：",
        ("🔹 标准模版 (自动读E/F/G列)", "🔧 自定义模式 (手动选3列)"),
        help="【标准模版】：适用于公司标准底稿（第5,6,7列为万元数据）。\n【自定义模式】：适用于任意格式表，由你指定哪三列是数据。"
    )
    
    st.markdown("---")
    uploaded_excel = st.file_uploader("Excel 底稿 (必须)", type=["xlsx", "xlsm"])
    uploaded_word_files = st.file_uploader("Word 附注 (可选)", type=["docx"], accept_multiple_files=True)
    header_row = st.number_input("表头所在行 (默认2)", value=2)
    st.markdown("### 3. Excel Sheet 匹配")
    sheet_asset = st.text_input("资产表 Sheet 名", value="1.合并资产表")
    sheet_liab = st.text_input("负债表 Sheet 名", value="2.合并负债表") 

# ================= 4. 主程序 =================

if not uploaded_excel:
    st.title("📊 财务分析报告自动化助手")
    st.markdown("""
    ### 💡 使用说明：
    1. **上传 Excel 底稿 (必须)**：请在左侧侧边栏上传。
    2. **上传 Word 附注 (可选)**：支持上传多个 Word 文件，用于生成原因分析。
    3. **选择读取模式**：
       - 如果是标准模版，直接用 **标准模版**。
       - 如果是普通表格，请切换到 **自定义模式** 并手动勾选三列数据。
    4. **一键导出**：支持导出 **精排版 Word 表格**，直接粘贴到报告中。
    """)
    st.info("👈 请先在左侧侧边栏上传 Excel 文件以开始使用。")

else:
    # Word 处理逻辑
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

    # 🔥 核心升级：交互式数据读取逻辑
    def get_clean_data(sheet_name):
        try:
            # 1. 先读全部数据
            df_full = pd.read_excel(uploaded_excel, sheet_name=sheet_name, header=header_row)
            
            # 2. 获取所有列名
            all_cols = df_full.columns.tolist()
            
            # 3. 确定数据列
            target_cols = []
            
            if "标准模版" in data_col_mode:
                # 默认读取 E, F, G (索引 4, 5, 6)
                if len(all_cols) > 6:
                    target_cols = [all_cols[0], all_cols[4], all_cols[5], all_cols[6]]
                else:
                    st.error("❌ 标准模版模式下，表格列数不足 7 列，请切换到【自定义模式】。")
                    return None, None, "列数不足"
            else:
                # 🔧 自定义模式：显示多选框让用户选
                st.info("👇 **【通用模式】请在下方选择 3 列包含数据的列**（请按顺序：最新一期 -> 上期 -> 上上期）：")
                
                # 排除第一列（通常是科目），让用户选数据列
                user_selected = st.multiselect(
                    "请勾选列（需选3个）：",
                    options=all_cols,
                    default=all_cols[1:4] if len(all_cols) >= 4 else None,
                    key=f"cols_{sheet_name}" # 避免Key冲突
                )
                
                if len(user_selected) != 3:
                    st.warning("⚠️ 请必须且只能选择 **3** 列数据！")
                    st.stop() # 暂停往下执行，等待用户选好
                
                # 拼装：[科目列] + [用户选的3列]
                # 注意：这里我们假设第一列永远是科目。
                # 为了防止用户把科目列也选进去了，我们强制使用 df_full.iloc[:, 0] 作为科目列
                df_subject = df_full.iloc[:, [0]]
                df_data = df_full[user_selected]
                
                # 合并
                df = pd.concat([df_subject, df_data], axis=1)
            
            if "标准模版" in data_col_mode:
                df = df_full.iloc[:, target_cols].copy()

            # 5. 提取日期标签
            orig_cols = df.columns.tolist()
            d_labels = [extract_date_label(orig_cols[1]), extract_date_label(orig_cols[2]), extract_date_label(orig_cols[3])]
            
            # 6. 标准化处理
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
        elif err != "列数不足": 
            st.error(f"❌ 读取 Excel 失败：{err}\n请检查【资产表 Sheet 名】是否为：{sheet_asset}")

    elif analysis_page == "(二) 负债结构分析":
        df_liab, d_labels, err = get_clean_data(sheet_liab)
        if df_liab is not None:
            total_name = "负债合计" 
            if not df_liab.index.str.contains(total_name).any():
                total_name = "负债总计"
            process_analysis_tab(df_liab, word_data_list, total_name, "负债", d_labels)
        elif err != "列数不足":
            st.error(f"❌ 读取 Excel 失败：{err}\n请检查【负债表 Sheet 名】是否为：{sheet_liab}")

    else:
        st.info("🚧 该模块正在施工中，敬请期待后续更新...")
