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

def create_word_table_file(df, title="数据表", bold_rows=None):
    """🔥 生成精排版 Word 表格 (五号字体，全居中)"""
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
    
    col_widths = [Cm(6.0)] + [Cm(3.0)] * (len(export_df.columns) - 1)
    for i, width in enumerate(col_widths):
        for row in table.rows:
            row.cells[i].width = width

    hdr_cells = table.rows[0].cells
    table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    table.rows[0].height = Cm(1.0)

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
            run.font.name = 'Times New Roman'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    for r_idx, row in export_df.iterrows():
        row_cells = table.add_row().cells
        table.rows[r_idx+1].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        table.rows[r_idx+1].height = Cm(0.8)
        subject_name = str(row[0]).strip()
        is_bold = False
        if bold_rows and subject_name in bold_rows: is_bold = True
        elif any(k in subject_name for k in ["合计", "总计", "净额", "净增加额", "构成", "活动"]): is_bold = True
        elif subject_name.endswith("：") or subject_name.endswith(":"): is_bold = True

        for i, val in enumerate(row):
            cell = row_cells[i]
            cell.text = str(val) if pd.notna(val) and val != "" else ""
            bottom_sz = 12 if r_idx == len(export_df) - 1 else 4
            set_cell_border(cell, top={"val": "single", "sz": 4}, bottom={"val": "single", "sz": bottom_sz}, left={"val": "single", "sz": 4}, right={"val": "single", "sz": 4})
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            paragraph = cell.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.size = Pt(10.5)
                run.font.name = 'Times New Roman'
                run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                if is_bold: run.font.bold = True
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
        full_text = []
        for p in doc.paragraphs:
            txt = p.text.strip()
            if len(txt) > 2: full_text.append(txt)
        for table in doc.tables:
            for row in table.rows:
                row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if row_text: full_text.append(" | ".join(row_text))
            full_text.append("\n")
        return "\n".join(full_text), True, ""
    except Exception as e:
        return "", False, f"❌ 读取失败: {str(e)}"

def find_context(subject, word_data_list):
    if not word_data_list: return ""
    clean_sub = subject.replace(" ", "")
    found_contexts = []
    for item in word_data_list:
        content = item['content']
        source = item['source']
        matches = list(re.finditer(re.escape(clean_sub), content))
        if matches:
            top_matches = matches[:3] 
            file_context = []
            for m in top_matches:
                idx = m.start()
                start = max(0, idx - 300)
                end = min(len(content), idx + 800)
                ctx = content[start:end].replace('\n', ' ')
                file_context.append(f"...{ctx}...")
            combined_ctx = "\n\n----------\n\n".join(file_context)
            found_contexts.append(f"📄 **来源：{source}**\n{combined_ctx}")
    return "\n\n====================\n\n".join(found_contexts)

def extract_date_label(header_str):
    s = str(header_str).strip()
    match = re.search(r'[【\[](.*?)[】\]]', s)
    if match: return match.group(1)
    year = re.search(r'(\d{4})', s)
    if year: return f"{year.group(1)}年"
    return s

def safe_pct(num, denom):
    return (num / denom * 100) if denom != 0 and pd.notna(num) and pd.notna(denom) else 0.0

def fuzzy_load_excel(file_obj, sheet_name, header_row=None):
    try:
        xl = pd.ExcelFile(file_obj)
        all_sheet_names = xl.sheet_names
        target_sheet = None
        
        # 1. 找 Sheet
        if sheet_name in all_sheet_names:
            target_sheet = sheet_name
        else:
            clean_target = sheet_name.replace(" ", "")
            for actual_name in all_sheet_names:
                if actual_name.replace(" ", "") == clean_target:
                    target_sheet = actual_name
                    st.toast(f"⚠️ 自动修正 Sheet 名为：'{actual_name}'")
                    break
        
        if target_sheet is None:
            return None, all_sheet_names

        # 2. 如果是财务指标表，使用智能读取逻辑 (忽略 header_row 参数)
        if "财务指标" in sheet_name or "5-3" in sheet_name:
            return smart_load_ratios(file_obj, target_sheet)
        
        # 3. 普通表读取
        return pd.read_excel(file_obj, sheet_name=target_sheet, header=header_row), None

    except Exception as e:
        return None, [str(e)]

# 🔥 核心升级：智能读取财务指标表 (自动找表头，自动找列)
def smart_load_ratios(file_obj, sheet_name):
    try:
        # 1. 先不带表头读取，扫描前 10 行找 "项目" 或 "指标"
        df_raw = pd.read_excel(file_obj, sheet_name=sheet_name, header=None)
        header_idx = -1
        for i in range(10):
            # 把这一行转成字符串，看是否包含关键词
            row_values = df_raw.iloc[i].astype(str).values
            if any("项目" in v or "指标" in v for v in row_values):
                header_idx = i
                break
        
        if header_idx == -1:
            # 兜底：如果找不到，就默认用第 2 行 (Index 1)
            header_idx = 1
            
        # 2. 重新读取，指定正确的表头行
        df = pd.read_excel(file_obj, sheet_name=sheet_name, header=header_idx)
        
        # 3. 智能寻找 T, T-1, T-2 所在的列
        # 规则：寻找包含 "年"、"期"、"T"、"20" 的列
        cols = df.columns.tolist()
        date_col_indices = []
        item_col_idx = 0 # 默认第一列是项目
        
        for idx, col_name in enumerate(cols):
            s = str(col_name)
            if "年" in s or "T" in s or "202" in s or "期" in s:
                date_col_indices.append(idx)
        
        # 如果找到至少 3 个日期列，取前 3 个
        if len(date_col_indices) >= 3:
            target_cols = [item_col_idx] + date_col_indices[:3]
        else:
            # 找不到就猜：通常是 第1列 + 第3,4,5列 (C,D,E) 或者是 (B,C,D)
            # 根据用户报错 "Unnamed: 4" (Col E)，推测可能是 C, D, E
            target_cols = [0, 2, 3, 4]
            
        df_final = df.iloc[:, target_cols]
        
        # 4. 标准化列名
        orig_cols = df_final.columns.tolist()
        d_labels = [extract_date_label(c) for c in orig_cols[1:]]
        df_final.columns = ['科目', 'T', 'T_1', 'T_2']
        
        # 5. 清洗数据
        df_final = df_final.dropna(subset=['科目'])
        df_final['科目'] = df_final['科目'].astype(str).str.strip()
        for c in ['T', 'T_1', 'T_2']:
            df_final[c] = pd.to_numeric(df_final[c], errors='coerce').fillna(0)
        df_final.set_index('科目', inplace=True)
        
        return df_final, d_labels

    except Exception as e:
        # 如果出错，抛出异常信息供调试
        raise Exception(f"智能读取失败: {str(e)}")

def find_row_fuzzy(df, keywords, default_val=None):
    if isinstance(keywords, str): keywords = [keywords]
    clean_index = df.index.astype(str).str.replace(r'\s+', '', regex=True)
    for kw in keywords:
        clean_kw = kw.replace(" ", "")
        mask = clean_index == clean_kw 
        if mask.any(): return df.loc[df.index[mask][0]]
    for kw in keywords:
        clean_kw = kw.replace(" ", "")
        mask = clean_index.str.contains(clean_kw, case=False, na=False)
        if mask.any(): return df.loc[df.index[mask][0]]
    if default_val is not None: return default_val
    return pd.Series(0, index=df.columns)

def find_index_fuzzy(df, keywords):
    if isinstance(keywords, str): keywords = [keywords]
    clean_index = df.index.astype(str).str.replace(r'\s+', '', regex=True)
    for kw in keywords:
        clean_kw = kw.replace(" ", "")
        mask = clean_index.str.contains(clean_kw, case=False, na=False)
        if mask.any(): return df.index.get_loc(df.index[mask][0])
    return None

# ================= 3. 业务逻辑：资产/负债 =================
def process_analysis_tab(df_raw, word_data_list, total_col_name, analysis_name, d_labels):
    try:
        if analysis_name == "负债":
             index_series = df_raw.index.astype(str)
             clean_index = index_series.str.replace(r'\s+', '', regex=True)
             clean_target = total_col_name.replace(" ", "")
             match_mask = (clean_index == clean_target)
             if match_mask.any():
                 target_label = df_raw.index[match_mask][0]
                 idx_pos = df_raw.index.get_loc(target_label)
                 if isinstance(idx_pos, slice): idx_pos = idx_pos.stop - 1
                 elif hasattr(idx_pos, '__iter__'): idx_pos = idx_pos[-1]
                 if isinstance(idx_pos, int): df_raw = df_raw.iloc[:idx_pos + 1]
        
        total_row = find_row_fuzzy(df_raw, [total_col_name])
        if total_row.sum() == 0 and total_row.name is None:
             st.error(f"❌ 未找到合计行：{total_col_name}")
             return
    except Exception as e:
        st.error(f"❌ 数据处理错误: {e}")
        return

    df = df_raw.copy()
    for period in ['T', 'T_1', 'T_2']:
        total = total_row[period]
        if total != 0: df[f'占比_{period}'] = df[period] / total
        else: df[f'占比_{period}'] = 0.0

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
                text = (f"报告期内，发行人负债总额分别为{total_row['T_2']:,.2f}万元、{total_row['T_1']:,.2f}万元和{total_row['T']:,.2f}万元。\n\n"
                        f"{d_labels[1]}较{d_labels[2]}{dir_prev}{abs(diff_prev):,.2f}万元，{label_prev}{abs(pct_prev):.2f}%；"
                        f"{d_labels[0]}发行人负债较{d_labels[1]}{dir_curr}{abs(diff_curr):,.2f}万元，{label_curr}{abs(pct_curr):.2f}%。"
                        f"报告期内发行人的负债规模呈现{trend_desc}态势。\n\n"
                        f"从负债结构来看，报告期内，流动负债分别为{curr_row['T_2']:,.2f}万元、{curr_row['T_1']:,.2f}万元和{curr_row['T']:,.2f}万元，"
                        f"占负债总额比例分别为{safe_pct(curr_row['T_2'], total_row['T_2']):.2f}%、{safe_pct(curr_row['T_1'], total_row['T_1']):.2f}%和{safe_pct(curr_row['T'], total_row['T']):.2f}%，"
                        f"主要由 **{'、'.join(top_5)}** 等构成；\n\n"
                        f"非流动负债分别为{non_curr_row['T_2']:,.2f}万元、{non_curr_row['T_1']:,.2f}万元和{non_curr_row['T']:,.2f}万元，"
                        f"占负债总额比例分别为{safe_pct(non_curr_row['T_2'], total_row['T_2']):.2f}%、{safe_pct(non_curr_row['T_1'], total_row['T_1']):.2f}%和{safe_pct(non_curr_row['T'], total_row['T']):.2f}%。")
            
            with st.container(border=True):
                st.markdown(f"#### 📝 {analysis_name}综述文案")
                st.text_area("文案内容", value=text, height=300, label_visibility="collapsed")
                st.caption("✨ 已自动优化排版，支持自动换行。点击框内按 Ctrl+A 即可全选。")

        except Exception as e:
             st.error(f"生成文案出错: {e}")

    with tab3:
        latest_date_label = d_labels[0]
        st.info(f"💡 **提示**：以下科目占比均基于 **{latest_date_label} (最新一期)** 的数据计算。")
        if word_data_list: st.success(f"✅ **附注加载成功**：已结合 **{len(word_data_list)} 个 Word 附注** 生成指令。")
        else: st.warning("⚠️ 未检测到 Word 附注，仅基于 Excel 数据生成指令。")
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
            prompt = f"""【任务】分析“{subject}”变动原因。\n\n【1. 数据趋势】\n{d_t2}、{d_t1}及{d_t}，发行人{subject}余额分别为{row['T_2']:,.2f}万元、{row['T_1']:,.2f}万元和{row['T']:,.2f}万元，占{denom_text}的比例分别为{row['占比_T_2']*100:.2f}%、{row['占比_T_1']*100:.2f}%和{row['占比_T']*100:.2f}%。\n\n【2. 变动情况】\n截至{d_t1}，发行人{subject}较{d_t2}{dir_prev}{abs(diff_prev):,.2f}万元，{label_prev}{abs(pct_prev):.2f}%；\n截至{d_t}，发行人{subject}较{d_t1}{dir_curr}{abs(diff_curr):,.2f}万元，{label_curr}{abs(pct_curr):.2f}%。"""
            if word_data_list: prompt += f"""\n\n【3. 附注线索】\n{find_context(subject, word_data_list)}\n\n【4. 写作要求】\n结合数据和附注分析原因。"""
            with st.expander(f"📌 {subject} (占比 {row['占比_T']:.2%} @ {latest_date_label})"):
                st.text_area(label="AI 指令", value=prompt, height=250, key=f"area_{subject}", label_visibility="collapsed")

# ================= 4. 业务逻辑：现金流量 =================
def calculate_cash_flow_percentages(df_raw, d_labels):
    data_list = []
    d_t, d_t1, d_t2 = d_labels
    sections = [
        (["经营活动产生的现金流量", "一、经营活动"], ["经营活动现金流入小计"], "一、经营活动现金流入构成"),
        (["经营活动现金流入小计"], ["经营活动现金流出小计"], "二、经营活动现金流出构成"),
        (["投资活动产生的现金流量", "二、投资活动"], ["投资活动现金流入小计"], "三、投资活动现金流入构成"),
        (["投资活动现金流入小计"], ["投资活动现金流出小计"], "四、投资活动现金流出构成"),
        (["筹资活动产生的现金流量", "三、筹资活动"], ["筹资活动现金流入小计"], "五、筹资活动现金流入构成"),
        (["筹资活动现金流入小计"], ["筹资活动现金流出小计"], "六、筹资活动现金流出构成"),
    ]
    for start_kws, end_kws, cat_name in sections:
        data_list.append([cat_name, "", "", ""])
        idx_start = find_index_fuzzy(df_raw, start_kws)
        idx_end = find_index_fuzzy(df_raw, end_kws)
        if idx_start is not None and idx_end is not None and idx_end > idx_start:
            denom_row = df_raw.iloc[idx_end]
            subset = df_raw.iloc[idx_start+1 : idx_end]
            for i in range(len(subset)):
                row = subset.iloc[i]
                subject = row.name
                if not isinstance(subject, str) or len(subject.strip()) < 2: continue
                pct_t = safe_pct(row['T'], denom_row['T'])
                pct_t1 = safe_pct(row['T_1'], denom_row['T_1'])
                pct_t2 = safe_pct(row['T_2'], denom_row['T_2'])
                data_list.append([subject, f"{pct_t:.2f}%", f"{pct_t1:.2f}%", f"{pct_t2:.2f}%"])
    return pd.DataFrame(data_list, columns=["项目", f"{d_t}占比", f"{d_t1}占比", f"{d_t2}占比"]).set_index("项目")

def process_cash_flow_tab(df_raw, word_data_list, d_labels):
    d_t, d_t1, d_t2 = d_labels
    
    structure = [("经营活动产生的现金流量：", None), ("经营活动现金流入小计", ["经营活动现金流入小计"]), ("经营活动现金流出小计", ["经营活动现金流出小计"]), ("经营活动产生的现金流量净额", ["经营活动产生的现金流量净额"]), ("投资活动产生的现金流量：", None), ("投资活动现金流入小计", ["投资活动现金流入小计"]), ("投资活动现金流出小计", ["投资活动现金流出小计"]), ("投资活动产生的现金流量净额", ["投资活动产生的现金流量净额"]), ("筹资活动产生的现金流量：", None), ("筹资活动现金流入小计", ["筹资活动现金流入小计"]), ("筹资活动现金流出小计", ["筹资活动现金流出小计"]), ("筹资活动产生的现金流量净额", ["筹资活动产生的现金流量净额"]), ("现金及现金等价物净增加额", ["现金及现金等价物净增加额"])]
    data_list = []
    for display_name, keywords in structure:
        if keywords is None: data_list.append([display_name, "", "", ""])
        else:
            row = find_row_fuzzy(df_raw, keywords)
            if row.name is None: val_t, val_t1, val_t2 = 0, 0, 0
            else: val_t, val_t1, val_t2 = row['T'], row['T_1'], row['T_2']
            data_list.append([display_name, f"{val_t:,.2f}" if val_t!="" else "", f"{val_t1:,.2f}" if val_t1!="" else "", f"{val_t2:,.2f}" if val_t2!="" else ""])
    df_display = pd.DataFrame(data_list, columns=["项目", d_t, d_t1, d_t2])
    df_display.set_index("项目", inplace=True)

    df_pct = calculate_cash_flow_percentages(df_raw, d_labels)

    tab1, tab2, tab3, tab4 = st.tabs(["📋 摘要数据", "📊 占比分析", "📝 综述文案", "🤖 AI 分析指令"])
    
    with tab1:
        c1, c2, c3 = st.columns([6, 1.2, 1.2]) 
        with c1: st.markdown("### 现金流量结构明细")
        with c2:
            doc_file = create_word_table_file(df_display, title="现金流量表摘要")
            st.download_button("📥 下载 Word", doc_file, "现金流量表.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        with c3:
            excel_file = create_excel_file(df_display)
            st.download_button("📥 下载 Excel", excel_file, "现金流量表.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(df_display, use_container_width=True)

    with tab2:
        c1, c2 = st.columns([6, 1.5])
        with c1: st.markdown("### 各项活动现金流占比分析")
        with c2:
            doc_pct = create_word_table_file(df_pct, title="现金流量占比表")
            st.download_button("📥 下载占比表 Word", doc_pct, "现金流占比.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        st.info("💡 说明：流入项占比 = 科目/流入小计；流出项占比 = 科目/流出小计")
        st.dataframe(df_pct, use_container_width=True)

    with tab3:
        op_in_total = find_row_fuzzy(df_raw, ["经营活动现金流入小计"])
        op_out_total = find_row_fuzzy(df_raw, ["经营活动现金流出小计"])
        op_net = find_row_fuzzy(df_raw, ["经营活动产生的现金流量净额"])
        op_sales = find_row_fuzzy(df_raw, ["销售商品、提供劳务收到的现金"])
        op_other_in = find_row_fuzzy(df_raw, ["收到其他与经营活动有关的现金"])
        op_buy = find_row_fuzzy(df_raw, ["购买商品、接受劳务支付的现金"])
        op_other_out = find_row_fuzzy(df_raw, ["支付其他与经营活动有关的现金"])
        inv_net = find_row_fuzzy(df_raw, ["投资活动产生的现金流量净额"])
        inv_in_total = find_row_fuzzy(df_raw, ["投资活动现金流入小计"])
        inv_out_total = find_row_fuzzy(df_raw, ["投资活动现金流出小计"])
        inv_buy_asset = find_row_fuzzy(df_raw, ["购建固定资产、无形资产和其他长期资产支付的现金"])
        fin_net = find_row_fuzzy(df_raw, ["筹资活动产生的现金流量净额"])
        fin_in_total = find_row_fuzzy(df_raw, ["筹资活动现金流入小计"])
        fin_borrow_in = find_row_fuzzy(df_raw, ["取得借款收到的现金"])
        fin_invest_in = find_row_fuzzy(df_raw, ["吸收投资收到的现金"])
        fin_out_total = find_row_fuzzy(df_raw, ["筹资活动现金流出小计"])
        fin_repay = find_row_fuzzy(df_raw, ["偿还债务支付的现金"])
        fin_interest = find_row_fuzzy(df_raw, ["分配股利、利润或偿付利息支付的现金"])

        # Box 1
        with st.container(border=True):
            st.markdown("#### 📝 1、经营活动产生的现金流量分析")
            text_op = (f"报告期内，发行人经营活动现金流入分别为{op_in_total['T_2']:,.2f}万元、{op_in_total['T_1']:,.2f}万元和{op_in_total['T']:,.2f}万元。\n\n"
                     f"其中，销售商品、提供劳务收到的现金分别为{op_sales['T_2']:,.2f}万元、{op_sales['T_1']:,.2f}万元及{op_sales['T']:,.2f}万元，"
                     f"占经营活动现金流入的{safe_pct(op_sales['T_2'], op_in_total['T_2']):.2f}%、{safe_pct(op_sales['T_1'], op_in_total['T_1']):.2f}%及{safe_pct(op_sales['T'], op_in_total['T']):.2f}%；\n\n"
                     f"收到其他与经营活动有关的现金分别为{op_other_in['T_2']:,.2f}万元、{op_other_in['T_1']:,.2f}万元及{op_other_in['T']:,.2f}万元，"
                     f"占经营活动现金流入的{safe_pct(op_other_in['T_2'], op_in_total['T_2']):.2f}%、{safe_pct(op_other_in['T_1'], op_in_total['T_1']):.2f}%及{safe_pct(op_other_in['T'], op_in_total['T']):.2f}%。"
                     f"发行人收到其他与经营活动有关的现金主要包括（）。\n\n")
            text_op += (f"报告期内，发行人经营活动现金流出分别为{op_out_total['T_2']:,.2f}万元、{op_out_total['T_1']:,.2f}万元和{op_out_total['T']:,.2f}万元。\n\n"
                     f"报告期内，发行人经营活动现金流出主要来源于购买商品、接受劳务支付的现金和支付其他与经营活动有关的现金。"
                     f"报告期内，发行人购买商品、接受劳务支付的现金分别为{op_buy['T_2']:,.2f}万元、{op_buy['T_1']:,.2f}万元及{op_buy['T']:,.2f}万元，"
                     f"占经营活动现金流出的{safe_pct(op_buy['T_2'], op_out_total['T_2']):.2f}%、{safe_pct(op_buy['T_1'], op_out_total['T_1']):.2f}%及{safe_pct(op_buy['T'], op_out_total['T']):.2f}%。\n\n"
                     f"发行人支付其他与经营活动有关的现金分别为{op_other_out['T_2']:,.2f}万元、{op_other_out['T_1']:,.2f}万元及{op_other_out['T']:,.2f}万元，"
                     f"占经营活动现金流出的{safe_pct(op_other_out['T_2'], op_out_total['T_2']):.2f}%、{safe_pct(op_other_out['T_1'], op_out_total['T_1']):.2f}%及{safe_pct(op_other_out['T'], op_out_total['T']):.2f}%。"
                     f"支付其他与经营活动有关的现金包括：（）等。\n\n")
            text_op += (f"报告期内，发行人经营活动产生的现金流量净额分别为{op_net['T_2']:,.2f}万元、{op_net['T_1']:,.2f}万元和{op_net['T']:,.2f}万元，"
                     f"主要系（）所致。")
            st.text_area("文案内容", value=text_op, height=350, label_visibility="collapsed", key="txt_op")

        # Box 2
        with st.container(border=True):
            st.markdown("#### 📝 2、投资活动产生的现金流量分析")
            text_inv = (f"报告期内，发行人投资活动产生的现金流量净额分别为{inv_net['T_2']:,.2f}万元、{inv_net['T_1']:,.2f}万元和{inv_net['T']:,.2f}万元。\n\n"
                     f"投资活动现金流入分别为{inv_in_total['T_2']:,.2f}万元、{inv_in_total['T_1']:,.2f}万元及{inv_in_total['T']:,.2f}万元；"
                     f"投资活动现金流出分别为{inv_out_total['T_2']:,.2f}万元、{inv_out_total['T_1']:,.2f}万元及{inv_out_total['T']:,.2f}万元，"
                     f"其中购建固定资产、无形资产和其他长期资产支付的现金分别为{inv_buy_asset['T_2']:,.2f}万元、{inv_buy_asset['T_1']:,.2f}万元及{inv_buy_asset['T']:,.2f}万元，"
                     f"占投资活动现金流出的{safe_pct(inv_buy_asset['T_2'], inv_out_total['T_2']):.2f}%、{safe_pct(inv_buy_asset['T_1'], inv_out_total['T_1']):.2f}%及{safe_pct(inv_buy_asset['T'], inv_out_total['T']):.2f}%。\n\n"
                     f"发行人投资活动现金流量净额（），主要是发行人（）所致。")
            st.text_area("文案内容", value=text_inv, height=250, label_visibility="collapsed", key="txt_inv")

        # Box 3
        with st.container(border=True):
            st.markdown("#### 📝 3、筹资活动产生的现金流量分析")
            text_fin = (f"报告期内，发行人筹资活动产生的现金流量净额分别为{fin_net['T_2']:,.2f}万元、{fin_net['T_1']:,.2f}万元和{fin_net['T']:,.2f}万元。\n\n"
                     f"报告期内筹资活动产生的现金流量净额较大，主要系吸收投资收到的现金及取得借款收到的现金流入所致。\n\n")
            text_fin += (f"筹资活动现金流入方面，发行人筹资活动现金流入主要由取得借款所收到的现金及吸收投资收到的现金构成。"
                     f"{d_t2}、{d_t1}及{d_t}，发行人筹资活动产生的现金流入分别为{fin_in_total['T_2']:,.2f}万元、{fin_in_total['T_1']:,.2f}万元及{fin_in_total['T']:,.2f}万元，"
                     f"其中取得借款收到的现金分别为{fin_borrow_in['T_2']:,.2f}万元、{fin_borrow_in['T_1']:,.2f}万元及{fin_borrow_in['T']:,.2f}万元；"
                     f"吸收投资收到的现金分别为{fin_invest_in['T_2']:,.2f}万元、{fin_invest_in['T_1']:,.2f}万元及{fin_invest_in['T']:,.2f}万元。\n\n")
            text_fin += (f"{d_t2}、{d_t1}及{d_t}，发行人筹资活动产生的现金流出分别为{fin_out_total['T_2']:,.2f}万元、{fin_out_total['T_1']:,.2f}万元和{fin_out_total['T']:,.2f}万元。"
                     f"发行人筹资活动现金流出主要由偿还债务所支付的现金及分配股利、利润或偿付利息支付的现金构成。"
                     f"其中报告期内，发行人偿还债务支付的现金分别为{fin_repay['T_2']:,.2f}万元、{fin_repay['T_1']:,.2f}万元和{fin_repay['T']:,.2f}万元，"
                     f"分配股利、利润或偿付利息所支付的现金分别为{fin_interest['T_2']:,.2f}万元、{fin_interest['T_1']:,.2f}万元和{fin_interest['T']:,.2f}万元。")
            st.text_area("文案内容", value=text_fin, height=350, label_visibility="collapsed", key="txt_fin")

    with tab4:
        st.info("💡 **提示**：现金流量分析侧重于三大活动净额变动。")
        if word_data_list: st.success(f"✅ **附注加载成功**：已结合 **{len(word_data_list)} 个 Word 附注** 生成指令。")
        else: st.warning("⚠️ 未检测到 Word 附注，仅基于 Excel 数据生成指令。")
        target_subjects = ["经营活动产生的现金流量净额", "投资活动产生的现金流量净额", "筹资活动产生的现金流量净额"]
        for subject in target_subjects:
            row = find_row_fuzzy(df_raw, [subject])
            if row.name is None: continue
            diff_prev = row['T_1'] - row['T_2']
            diff_curr = row['T'] - row['T_1']
            dir_prev = "增加" if diff_prev >= 0 else "减少"
            dir_curr = "增加" if diff_curr >= 0 else "减少"
            prompt = f"""【任务】分析“{subject}”变动原因。\n\n【1. 数据趋势】\n{d_t2}、{d_t1}及{d_t}，发行人{subject}分别为{row['T_2']:,.2f}万元、{row['T_1']:,.2f}万元和{row['T']:,.2f}万元。\n\n【2. 变动情况】\n截至{d_t1}，较{d_t2}{dir_prev}{abs(diff_prev):,.2f}万元；\n截至{d_t}，较{d_t1}{dir_curr}{abs(diff_curr):,.2f}万元。"""
            if word_data_list: prompt += f"""\n\n【3. 附注线索】\n{find_context(subject, word_data_list)}\n\n【4. 写作要求】\n结合数据和附注分析原因。"""
            with st.expander(f"📌 {subject}"):
                st.text_area(label="AI 指令", value=prompt, height=200, key=f"cf_prompt_{subject}", label_visibility="collapsed")

# ================= 5. 业务逻辑：财务指标分析 =================
def process_financial_ratios_tab(df_raw, word_data_list, d_labels):
    d_t, d_t1, d_t2 = d_labels
    
    targets = [
        "资产负债率（%）",
        "流动比率（倍）",
        "速动比率（倍）",
        "EBITDA（万元）",
        "EBITDA利息保障倍数（倍）"
    ]
    
    data_list = []
    data_map = {} 
    
    for t in targets:
        keywords = [t, t.replace("（", "(").replace("）", ")")]
        row = find_row_fuzzy(df_raw, keywords)
        
        val_t, val_t1, val_t2 = 0, 0, 0
        if row.name is not None:
            val_t, val_t1, val_t2 = row['T'], row['T_1'], row['T_2']
            data_map[t] = {'T': val_t, 'T_1': val_t1, 'T_2': val_t2}
        
        if "EBITDA（万元）" in t:
            fmt_t = f"{val_t:,.2f}"
            fmt_t1 = f"{val_t1:,.2f}"
            fmt_t2 = f"{val_t2:,.2f}"
        else:
            fmt_t = f"{val_t:.2f}"
            fmt_t1 = f"{val_t1:.2f}"
            fmt_t2 = f"{val_t2:.2f}"
            
        data_list.append([t, fmt_t, fmt_t1, fmt_t2])

    df_display = pd.DataFrame(data_list, columns=["项目", d_t, d_t1, d_t2])
    df_display.set_index("项目", inplace=True)

    tab1, tab2, tab3 = st.tabs(["📋 指标数据", "📝 综述文案", "🤖 AI 分析指令"])

    with tab1:
        c1, c2, c3 = st.columns([6, 1.2, 1.2]) 
        with c1: st.markdown("### 主要偿债指标")
        with c2:
            doc_file = create_word_table_file(df_display, title="主要财务指标表")
            st.download_button("📥 下载 Word", doc_file, "财务指标表.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        with c3:
            excel_file = create_excel_file(df_display)
            st.download_button("📥 下载 Excel", excel_file, "财务指标表.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(df_display, use_container_width=True)

    with tab2:
        alr = data_map.get("资产负债率（%）", {'T':0,'T_1':0,'T_2':0})
        cr = data_map.get("流动比率（倍）", {'T':0,'T_1':0,'T_2':0})
        qr = data_map.get("速动比率（倍）", {'T':0,'T_1':0,'T_2':0})
        ebitda = data_map.get("EBITDA（万元）", {'T':0,'T_1':0,'T_2':0})
        int_cov = data_map.get("EBITDA利息保障倍数（倍）", {'T':0,'T_1':0,'T_2':0})

        with st.container(border=True):
            st.markdown("#### 📝 偿债能力分析综述")
            
            text = f"1、资产负债率\n\n"
            text += f"报告期内，发行人的资产负债率分别为{alr['T_2']:.2f}%、{alr['T_1']:.2f}%和{alr['T']:.2f}%。\n\n"
            
            text += f"2、流动比率及速动比率\n\n"
            text += (f"报告期内，发行人的流动比率分别为{cr['T_2']:.2f}倍、{cr['T_1']:.2f}倍和{cr['T']:.2f}倍；"
                     f"同期速动比率分别为{qr['T_2']:.2f}倍、{qr['T_1']:.2f}倍和{qr['T']:.2f}倍。\n\n")
            
            text += f"3、EBITDA利息保障倍数\n\n"
            text += (f"报告期内，发行人EBITDA分别为{ebitda['T_2']:,.2f}万元、{ebitda['T_1']:,.2f}万元和{ebitda['T']:,.2f}万元，"
                     f"发行人EBITDA利息保障倍数分别为{int_cov['T_2']:.2f}倍、{int_cov['T_1']:.2f}倍和{int_cov['T']:.2f}倍。")
            
            st.text_area("文案内容", value=text, height=400, label_visibility="collapsed", help="Ctrl+A 全选")

    with tab3:
        st.info("💡 **提示**：财务指标的变动通常需要结合资产负债结构和盈利能力进行综合分析。")
        prompts = [
            ("资产负债率", alr, "分析偿债风险变化"),
            ("流动比率", cr, "分析短期偿债能力"),
            ("EBITDA", ebitda, "分析盈利及获现能力")
        ]
        for name, data, task in prompts:
            prompt = f"""【任务】{task}：{name}\n\n【数据表现】\n{d_t2}、{d_t1}及{d_t}，{name}分别为{data['T_2']:.2f}、{data['T_1']:.2f}和{data['T']:.2f}。\n\n【分析要求】\n结合企业经营情况，分析指标变动的原因。"""
            with st.expander(f"📌 {name}"):
                st.text_area("AI 指令", value=prompt, height=200, label_visibility="collapsed")


# ================= 3. 侧边栏 =================
with st.sidebar:
    st.title("🎛️ 操控台")
    analysis_page = st.radio("请选择要生成的章节：", ["(一) 资产结构分析", "(二) 负债结构分析", "(三) 现金流量分析", "(四) 财务指标分析"])
    st.markdown("---")
    
    uploaded_excel = st.file_uploader("Excel 底稿 (必须)", type=["xlsx", "xlsm"])
    uploaded_word_files = st.file_uploader("Word 附注 (可选)", type=["docx"], accept_multiple_files=True)
    
    with st.expander("⚙️ 高级设置 (Sheet名称/表头行)"):
        header_row = st.number_input("表头所在行 (默认2，即第3行)", value=2, min_value=0)
        sheet_asset = st.text_input("资产表 Sheet 名", value="1.合并资产表")
        sheet_liab = st.text_input("负债表 Sheet 名", value="2.合并负债及权益表")
        sheet_cash = st.text_input("现金流量表 Sheet 名", value="4.合并现金流量表")
        sheet_ratios = st.text_input("财务指标表 Sheet 名", value="5-3主要财务指标计算-方案3（专用公司债）")

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
        * 现金流 -> `4.合并现金流量表`
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
    if uploaded_word_files:
        for w in uploaded_word_files:
            content, success, err_msg = load_single_word(w) 
            if success: word_data_list.append({'source': w.name, 'content': content})
            else: st.error(err_msg)
    if uploaded_word_files and not word_data_list: st.stop()
    elif uploaded_word_files: st.success(f"✅ **附注加载成功**：已读取 {len(word_data_list)} 个 Word 文件")

    def get_clean_data(target_sheet_name):
        try:
            df, all_sheets_if_failed = fuzzy_load_excel(uploaded_excel, target_sheet_name, header_row)
            if df is None: return None, None, f"未找到 Sheet '{target_sheet_name}' (现有 Sheet: {all_sheets_if_failed})"
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
        except Exception as e: return None, None, str(e)

    st.header(f"📊 {analysis_page}")

    if analysis_page == "(一) 资产结构分析":
        df_asset, d_labels, err = get_clean_data(sheet_asset)
        if df_asset is not None: process_analysis_tab(df_asset, word_data_list, "资产总计", "资产", d_labels)
        else: st.error(f"❌ 读取失败：{err}")

    elif analysis_page == "(二) 负债结构分析":
        df_liab, d_labels, err = get_clean_data(sheet_liab)
        if df_liab is not None:
            total_name = "负债合计" 
            if not df_liab.index.str.contains(total_name).any(): total_name = "负债总计"
            process_analysis_tab(df_liab, word_data_list, total_name, "负债", d_labels)
        else: st.error(f"❌ 读取失败：{err}")

    elif analysis_page == "(三) 现金流量分析":
        df_cash, d_labels, err = get_clean_data(sheet_cash)
        if df_cash is not None:
            process_cash_flow_tab(df_cash, word_data_list, d_labels)
        else: st.error(f"❌ 读取失败：{err}")

    elif analysis_page == "(四) 财务指标分析":
        # 🔥 使用智能读取器，不依赖 get_clean_data
        df_ratios, d_labels = fuzzy_load_excel(uploaded_excel, sheet_ratios, header_row) # 这里 header_row 其实被忽略了
        if df_ratios is not None:
            process_financial_ratios_tab(df_ratios, word_data_list, d_labels)
        else: 
            st.error(f"❌ 读取失败：未找到 Sheet '{sheet_ratios}'")

    else:
        st.info("🚧 该模块正在施工中，敬请期待后续更新...")
