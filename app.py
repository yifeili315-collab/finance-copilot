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

        # 财务指标表特供逻辑
        if "财务指标" in sheet_name or "5-3" in sheet_name:
            return smart_load_ratios(file_obj, target_sheet)
        
        return pd.read_excel(file_obj, sheet_name=target_sheet, header=header_row), None

    except Exception as e:
        return None, [str(e)]

def smart_load_ratios(file_obj, sheet_name):
    try:
        df_raw = pd.read_excel(file_obj, sheet_name=sheet_name, header=None)
        header_idx = -1
        for i in range(10):
            row_values = df_raw.iloc[i].astype(str).values
            if any("项目" in v or "指标" in v for v in row_values):
                header_idx = i
                break
        if header_idx == -1: header_idx = 1
        df = pd.read_excel(file_obj, sheet_name=sheet_name, header=header_idx)
        cols = df.columns.tolist()
        date_col_indices = []
        for idx, col_name in enumerate(cols):
            s = str(col_name)
            if "年" in s or "T" in s or "202" in s or "期" in s:
                date_col_indices.append(idx)
        if len(date_col_indices) >= 3:
            target_cols = [0] + date_col_indices[:3]
        else:
            target_cols = [0, 2, 3, 4]
        df_final = df.iloc[:, target_cols]
        orig_cols = df_final.columns.tolist()
        d_labels = [extract_date_label(c) for c in orig_cols[1:]]
        df_final.columns = ['科目', 'T', 'T_1', 'T_2']
        df_final = df_final.dropna(subset=['科目'])
        df_final['科目'] = df_final['科目'].astype(str).str.strip()
        for c in ['T', 'T_1', 'T_2']:
            df_final[c] = pd.to_numeric(df_final[c], errors='coerce').fillna(0)
        df_final.set_index('科目', inplace=True)
        return df_final, d_labels
    except Exception as e:
        raise Exception(f"智能读取失败: {str(e)}")

def find_row_fuzzy(df, keywords, exclude_keywords=None, default_val=None):
    if isinstance(keywords, str): keywords = [keywords]
    clean_index = df.index.astype(str).str.replace(r'\s+', '', regex=True)
    found_rows = []
    for kw in keywords:
        clean_kw = kw.replace(" ", "")
        mask_exact = clean_index == clean_kw
        mask_contains = clean_index.str.contains(clean_kw, case=False, na=False)
        if exclude_keywords:
            for ex_kw in exclude_keywords:
                clean_ex = ex_kw.replace(" ", "")
                mask_contains = mask_contains & (~clean_index.str.contains(clean_ex, case=False, na=False))
        matched_indices = df.index[mask_exact | mask_contains].tolist()
        for idx in matched_indices:
            row = df.loc[idx]
            if isinstance(row, pd.DataFrame):
                for _, r in row.iterrows(): found_rows.append(r)
            else:
                found_rows.append(row)
    best_row = None
    max_non_zeros = -1
    for row in found_rows:
        non_zeros = 0
        if row['T'] != 0 and pd.notna(row['T']): non_zeros += 1
        if row['T_1'] != 0 and pd.notna(row['T_1']): non_zeros += 1
        if row['T_2'] != 0 and pd.notna(row['T_2']): non_zeros += 1
        if non_zeros > max_non_zeros:
            max_non_zeros = non_zeros
            best_row = row
    if best_row is not None: return best_row
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

def smart_scale_convert(val, subject_name="", is_ebitda=False, is_ratio=False):
    if pd.isna(val) or val == 0: return 0.0
    if "亿元" in subject_name: return val * 10000.0
    if "万元" in subject_name: return val
    if "元" in subject_name and "万元" not in subject_name and "亿元" not in subject_name: return val / 10000.0
    if is_ebitda:
        if abs(val) > 1000000: return val / 10000.0
        else: return val
    if is_ratio:
        if abs(val) < 1.0: return val * 100.0
        return val
    return val

# ================= 3. 业务逻辑 =================
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

    tab1, tab2, tab3 = st.tabs(["📋 明细数据", "📝 综述文案", "📝 变动分析文案"])

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
        
        # 🟢 清空以冒号结尾的标题行数据（如“流动资产：”）
        for idx in final_df.index:
            if str(idx).strip().endswith("：") or str(idx).strip().endswith(":"):
                final_df.loc[idx] = ""

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
                        f"报告期内发行人的负债规模呈现{trend_desc}态势，主要原因为发行人（用户自行分析）。\n\n"
                        f"从负债结构来看，报告期内，流动负债分别为{curr_row['T_2']:,.2f}万元、{curr_row['T_1']:,.2f}万元和{curr_row['T']:,.2f}万元，"
                        f"占负债总额比例分别为{safe_pct(curr_row['T_2'], total_row['T_2']):.2f}%、{safe_pct(curr_row['T_1'], total_row['T_1']):.2f}%和{safe_pct(curr_row['T'], total_row['T']):.2f}%，"
                        f"主要由 **{'、'.join(top_5)}** 等构成；\n\n"
                        f"非流动负债分别为{non_curr_row['T_2']:,.2f}万元、{non_curr_row['T_1']:,.2f}万元和{non_curr_row['T']:,.2f}万元，"
                        f"占负债总额比例分别为{safe_pct(non_curr_row['T_2'], total_row['T_2']):.2f}%、{safe_pct(non_curr_row['T_1'], total_row['T_1']):.2f}%和{safe_pct(non_curr_row['T'], total_row['T']):.2f}%。")
            
            with st.container(border=True):
                st.markdown(f"#### 📝 {analysis_name}综述文案")
                st.markdown(text)
                st.code(text, language='text')

        except Exception as e:
             st.error(f"生成文案出错: {e}")

    with tab3:
        latest_date_label = d_labels[0]
        st.info(f"💡 **提示**：已根据数据生成科目变动分析文案草稿。")
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
            
            # 生成变动分析文案
            analysis_text = (f"报告期各期末，发行人{subject}余额分别为{row['T_2']:,.2f}万元、{row['T_1']:,.2f}万元和{row['T']:,.2f}万元，"
                           f"占{denom_text}的比例分别为{row['占比_T_2']*100:.2f}%、{row['占比_T_1']*100:.2f}%和{row['占比_T']*100:.2f}%。\n\n"
                           f"{d_t1}末，发行人{subject}较{d_t2}末{dir_prev}{abs(diff_prev):,.2f}万元，{label_prev}{abs(pct_prev):.2f}%；"
                           f"{d_t}末，发行人{subject}较{d_t1}末{dir_curr}{abs(diff_curr):,.2f}万元，{label_curr}{abs(pct_curr):.2f}%。\n\n"
                           f"变动主要原因为：（请在此处补充具体的业务原因，例如：业务规模扩大/缩减、新增/偿还款项等）。")
            
            # 如果有附注上下文，展示在下方供参考
            ctx = find_context(subject, word_data_list)
            if ctx:
                analysis_text += f"\n\n【参考附注信息】\n{ctx}"

            with st.expander(f"📌 {subject} (占比 {row['占比_T']:.2%} @ {latest_date_label})"):
                st.markdown(analysis_text)
                st.code(analysis_text, language='text')

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

    tab1, tab2, tab3, tab4 = st.tabs(["📋 摘要数据", "📊 占比分析", "📝 综述文案", "📝 变动分析文案"])
    
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

        # 🟢 [修改]：改用 container + markdown + code
        with st.container(border=True):
            st.markdown("#### 📝 1、经营活动产生的现金流量分析")
            text_op = (f"报告期内，发行人经营活动现金流入分别为{op_in_total['T_2']:,.2f}万元、{op_in_total['T_1']:,.2f}万元和{op_in_total['T']:,.2f}万元。\n\n"
                     f"其中，销售商品、提供劳务收到的现金分别为{op_sales['T_2']:,.2f}万元、{op_sales['T_1']:,.2f}万元及{op_sales['T']:,.2f}万元，"
                     f"占经营活动现金流入的{safe_pct(op_sales['T_2'], op_in_total['T_2']):.2f}%、{safe_pct(op_sales['T_1'], op_in_total['T_1']):.2f}%及{safe_pct(op_sales['T'], op_in_total['T']):.2f}%；\n\n"
                     f"收到其他与经营活动有关的现金分别为{op_other_in['T_2']:,.2f}万元、{op_other_in['T_1']:,.2f}万元及{op_other_in['T']:,.2f}万元，"
                     f"占经营活动现金流入的{safe_pct(op_other_in['T_2'], op_in_total['T_2']):.2f}%、{safe_pct(op_other_in['T_1'], op_in_total['T_1']):.2f}%及{safe_pct(op_other_in['T'], op_in_total['T']):.2f}%。"
                     f"发行人收到其他与经营活动有关的现金主要包括【】。\n\n")
            text_op += (f"报告期内，发行人经营活动现金流出分别为{op_out_total['T_2']:,.2f}万元、{op_out_total['T_1']:,.2f}万元和{op_out_total['T']:,.2f}万元。\n\n"
                      f"报告期内，发行人经营活动现金流出主要来源于【】。"
                      f"报告期内，发行人购买商品、接受劳务支付的现金分别为{op_buy['T_2']:,.2f}万元、{op_buy['T_1']:,.2f}万元及{op_buy['T']:,.2f}万元，"
                      f"占经营活动现金流出的{safe_pct(op_buy['T_2'], op_out_total['T_2']):.2f}%、{safe_pct(op_buy['T_1'], op_out_total['T_1']):.2f}%及{safe_pct(op_buy['T'], op_out_total['T']):.2f}%。\n\n"
                      f"发行人支付其他与经营活动有关的现金分别为{op_other_out['T_2']:,.2f}万元、{op_other_out['T_1']:,.2f}万元及{op_other_out['T']:,.2f}万元，"
                      f"占经营活动现金流出的{safe_pct(op_other_out['T_2'], op_out_total['T_2']):.2f}%、{safe_pct(op_other_out['T_1'], op_out_total['T_1']):.2f}%及{safe_pct(op_other_out['T'], op_out_total['T']):.2f}%。"
                      f"支付其他与经营活动有关的现金包括：【】。\n\n")
            text_op += (f"报告期内，发行人经营活动产生的现金流量净额分别为{op_net['T_2']:,.2f}万元、{op_net['T_1']:,.2f}万元和{op_net['T']:,.2f}万元，"
                      f"主要系【】所致。")
            st.markdown(text_op)
            st.code(text_op, language='text')

        with st.container(border=True):
            st.markdown("#### 📝 2、投资活动产生的现金流量分析")
            text_inv = (f"报告期内，发行人投资活动产生的现金流量净额分别为{inv_net['T_2']:,.2f}万元、{inv_net['T_1']:,.2f}万元和{inv_net['T']:,.2f}万元。\n\n"
                      f"投资活动现金流入分别为{inv_in_total['T_2']:,.2f}万元、{inv_in_total['T_1']:,.2f}万元及{inv_in_total['T']:,.2f}万元；"
                      f"投资活动现金流出分别为{inv_out_total['T_2']:,.2f}万元、{inv_out_total['T_1']:,.2f}万元及{inv_out_total['T']:,.2f}万元，"
                      f"其中购建固定资产、无形资产和其他长期资产支付的现金分别为{inv_buy_asset['T_2']:,.2f}万元、{inv_buy_asset['T_1']:,.2f}万元及{inv_buy_asset['T']:,.2f}万元，"
                      f"占投资活动现金流出的{safe_pct(inv_buy_asset['T_2'], inv_out_total['T_2']):.2f}%、{safe_pct(inv_buy_asset['T_1'], inv_out_total['T_1']):.2f}%及{safe_pct(inv_buy_asset['T'], inv_out_total['T']):.2f}%。\n\n"
                      f"发行人投资活动现金流量净额【】，主要是发行人【】所致。")
            st.markdown(text_inv)
            st.code(text_inv, language='text')

        with st.container(border=True):
            st.markdown("#### 📝 3、筹资活动产生的现金流量分析")
            text_fin = (f"报告期内，发行人筹资活动产生的现金流量净额分别为{fin_net['T_2']:,.2f}万元、{fin_net['T_1']:,.2f}万元和{fin_net['T']:,.2f}万元。\n\n"
                      f"报告期内筹资活动产生的现金流量净额【】，主要系【】所致。\n\n")
            text_fin += (f"筹资活动现金流入方面，发行人筹资活动现金流入主要由【】构成。"
                       f"{d_t2}、{d_t1}及{d_t}，发行人筹资活动产生的现金流入分别为{fin_in_total['T_2']:,.2f}万元、{fin_in_total['T_1']:,.2f}万元及{fin_in_total['T']:,.2f}万元，"
                       f"其中取得借款收到的现金分别为{fin_borrow_in['T_2']:,.2f}万元、{fin_borrow_in['T_1']:,.2f}万元及{fin_borrow_in['T']:,.2f}万元；"
                       f"吸收投资收到的现金分别为{fin_invest_in['T_2']:,.2f}万元、{fin_invest_in['T_1']:,.2f}万元及{fin_invest_in['T']:,.2f}万元。\n\n")
            text_fin += (f"{d_t2}、{d_t1}及{d_t}，发行人筹资活动产生的现金流出分别为{fin_out_total['T_2']:,.2f}万元、{fin_out_total['T_1']:,.2f}万元和{fin_out_total['T']:,.2f}万元。"
                       f"发行人筹资活动现金流出主要由【】构成。"
                       f"其中报告期内，发行人偿还债务支付的现金分别为{fin_repay['T_2']:,.2f}万元、{fin_repay['T_1']:,.2f}万元和{fin_repay['T']:,.2f}万元，"
                       f"分配股利、利润或偿付利息所支付的现金分别为{fin_interest['T_2']:,.2f}万元、{fin_interest['T_1']:,.2f}万元和{fin_interest['T']:,.2f}万元。")
            st.markdown(text_fin)
            st.code(text_fin, language='text')

    with tab4:
        st.info("💡 **提示**：已自动生成净现金流量变动分析文案草稿。")
        target_subjects = ["经营活动产生的现金流量净额", "投资活动产生的现金流量净额", "筹资活动产生的现金流量净额"]
        for subject in target_subjects:
            row = find_row_fuzzy(df_raw, [subject])
            if row.name is None: continue
            diff_prev = row['T_1'] - row['T_2']
            diff_curr = row['T'] - row['T_1']
            dir_prev = "增加" if diff_prev >= 0 else "减少"
            dir_curr = "增加" if diff_curr >= 0 else "减少"
            
            # 生成变动分析文案
            cf_text = (f"报告期各期，发行人{subject}分别为{row['T_2']:,.2f}万元、{row['T_1']:,.2f}万元和{row['T']:,.2f}万元。\n\n"
                     f"{d_t1}，发行人{subject}较{d_t2}净{dir_prev}{abs(diff_prev):,.2f}万元；\n"
                     f"{d_t}，发行人{subject}较{d_t1}净{dir_curr}{abs(diff_curr):,.2f}万元。\n\n"
                     f"变动主要原因为：（请在此处补充具体的业务或资金变动原因）。")
            
            # 🟢 [修改]：Expander内直接展示 markdown + 代码框
            with st.expander(f"📌 {subject}"):
                st.markdown(cf_text)
                st.code(cf_text, language='text')

# ================= 5. 业务逻辑：盈利能力分析 (NEW!) =================
def process_profitability_tab(df_raw, word_data_list, d_labels):
    d_t, d_t1, d_t2 = d_labels
    
    # 1. 定义标准化的科目名称顺序
    standard_items = [
        "营业收入", "营业成本", "销售费用", "管理费用", "研发费用", "财务费用",
        "其他收益", "营业利润", "营业外收入", "营业外支出", "利润总额", "净利润",
        "营业毛利率", "平均总资产回报率"
    ]

    # 2. 查找关键数据行 (使用更灵活的模糊匹配)
    def get_row_data(keywords, default_zero=True):
        row = find_row_fuzzy(df_raw, keywords)
        if row.name:
            return row['T'], row['T_1'], row['T_2']
        return 0, 0, 0 if default_zero else (None, None, None)

    # 提取基础数据用于后续计算
    rev_t, rev_t1, rev_t2 = get_row_data(['营业收入'])
    cost_t, cost_t1, cost_t2 = get_row_data(['营业成本'])

    # 构建表格数据列表
    data_list = []
    
    for item in standard_items:
        # 特殊计算行
        if item == "营业毛利率":
            m_t = (rev_t - cost_t) / rev_t * 100 if rev_t != 0 else 0.0
            m_t1 = (rev_t1 - cost_t1) / rev_t1 * 100 if rev_t1 != 0 else 0.0
            m_t2 = (rev_t2 - cost_t2) / rev_t2 * 100 if rev_t2 != 0 else 0.0
            data_list.append([item, f"{m_t:.2f}", f"{m_t1:.2f}", f"{m_t2:.2f}"])
        elif item == "平均总资产回报率":
            # 暂无数据，留空
            data_list.append([item, "", "", ""])
        else:
            # 常规科目查找
            # 针对一些科目定义别名列表以提高命中率
            search_kws = [item]
            if item == "营业利润": search_kws = ['营业利润', '三、营业利润']
            elif item == "利润总额": search_kws = ['利润总额', '四、利润总额']
            elif item == "净利润": search_kws = ['净利润', '五、净利润']
            elif item == "研发费用": search_kws = ['研发费用'] # 确保能找到研发
            
            val_t, val_t1, val_t2 = get_row_data(search_kws)
            
            # 格式化
            f_t = f"{val_t:,.2f}" if val_t != 0 else "0.00"
            f_t1 = f"{val_t1:,.2f}" if val_t1 != 0 else "0.00"
            f_t2 = f"{val_t2:,.2f}" if val_t2 != 0 else "0.00"
            
            # 如果是其他收益且为0，可能想留空？这里统一显示0.00保持一致，或者根据需求改
            if item == "其他收益" and val_t == 0 and val_t1 == 0 and val_t2 == 0:
                 f_t, f_t1, f_t2 = "", "", ""

            data_list.append([item, f_t, f_t1, f_t2])

    # 转 DataFrame
    df_fmt = pd.DataFrame(data_list, columns=["项目", d_t, d_t1, d_t2])
    df_fmt.set_index("项目", inplace=True)

    # 4. 计算逻辑 (用于文案) - 重新获取一次以便文案生成使用方便
    margins = {
        'T': (rev_t - cost_t) / rev_t * 100 if rev_t != 0 else 0.0,
        'T_1': (rev_t1 - cost_t1) / rev_t1 * 100 if rev_t1 != 0 else 0.0,
        'T_2': (rev_t2 - cost_t2) / rev_t2 * 100 if rev_t2 != 0 else 0.0
    }
    
    # 重新计算期间费用总额 (文案用)
    def get_val(name):
        r = get_row_data([name])
        return {'T': r[0], 'T_1': r[1], 'T_2': r[2]}
        
    exp_items = ['销售费用', '管理费用', '研发费用', '财务费用']
    period_expenses = {'T': 0, 'T_1': 0, 'T_2': 0}
    for ex in exp_items:
        vals = get_val(ex)
        for k in period_expenses: period_expenses[k] += vals[k]

    pe_ratios = {}
    for col in ['T', 'T_1', 'T_2']:
        r_val = rev_t if col == 'T' else (rev_t1 if col == 'T_1' else rev_t2)
        pe_ratios[col] = period_expenses[col] / r_val * 100 if r_val != 0 else 0.0
    
    # 🟢 [新增]：查找期间费用分析所需的所有费用行
    idx_start = find_index_fuzzy(df_raw, ['营业总成本', '二、营业总成本'])
    idx_end = find_index_fuzzy(df_raw, ['资产减值损失', '加：资产减值损失', '投资收益'])
    
    all_expense_rows = []
    if idx_start and idx_end and idx_end > idx_start:
        subset = df_raw.iloc[idx_start+1 : idx_end]
        for i in range(len(subset)):
            row = subset.iloc[i]
            if "费用" in str(row.name):
                all_expense_rows.append(row)
    else:
        # Fallback if structure not found
        for kw in exp_items:
             r = find_row_fuzzy(df_raw, [kw])
             if r.name: all_expense_rows.append(r)

    # 🟢 [新增]：构建期间费用分析表格数据
    period_exp_data = []
    for r in all_expense_rows:
        row_dat = [r.name]
        
        # T (Latest)
        val_t = r['T']
        pct_t = val_t / rev_t * 100 if rev_t else 0
        row_dat.extend([f"{val_t:,.2f}", f"{pct_t:.2f}%"])
        
        # T-1
        val_t1 = r['T_1']
        pct_t1 = val_t1 / rev_t1 * 100 if rev_t1 else 0
        row_dat.extend([f"{val_t1:,.2f}", f"{pct_t1:.2f}%"])

        # T-2
        val_t2 = r['T_2']
        pct_t2 = val_t2 / rev_t2 * 100 if rev_t2 else 0
        row_dat.extend([f"{val_t2:,.2f}", f"{pct_t2:.2f}%"])
        
        period_exp_data.append(row_dat)
    
    # 🟢 [修复]：使用带年份的列名，避免 Duplicate column names 错误
    pe_cols = ["项目", 
               f"{d_t}金额", f"{d_t}占比", 
               f"{d_t1}金额", f"{d_t1}占比",
               f"{d_t2}金额", f"{d_t2}占比"]
    
    df_period_exp = pd.DataFrame(period_exp_data, columns=pe_cols).set_index("项目")

    # UI 展示
    # 🟢 [修改]：新增 Tab 2 期间费用分析
    tab1, tab2, tab3, tab4 = st.tabs(["📋 盈利能力明细", "📊 期间费用分析", "📝 综述文案", "📝 变动分析文案"])

    with tab1:
        c1, c2, c3 = st.columns([6, 1.2, 1.2]) 
        with c1: st.markdown("### 盈利能力明细表")
        with c2:
            doc_file = create_word_table_file(df_fmt, title="盈利能力分析表")
            st.download_button("📥 下载 Word", doc_file, "盈利能力表.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        with c3:
            excel_file = create_excel_file(df_fmt)
            st.download_button("📥 下载 Excel", excel_file, "盈利能力表.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(df_fmt, use_container_width=True)
    
    # 🟢 [新增]：期间费用分析 Tab 内容
    with tab2:
        c1, c2, c3 = st.columns([6, 1.2, 1.2])
        with c1: st.markdown("### 期间费用分析表")
        with c2:
            doc_file_pe = create_word_table_file(df_period_exp, title="期间费用分析表")
            st.download_button("📥 下载 Word", doc_file_pe, "期间费用分析表.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        with c3:
            excel_file_pe = create_excel_file(df_period_exp)
            st.download_button("📥 下载 Excel", excel_file_pe, "期间费用分析表.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(df_period_exp, use_container_width=True)

    with tab3:
        # 🟢 [修改]：改用 container + markdown + code
        with st.container(border=True):
            st.markdown("#### 📝 1、营业收入、营业成本和毛利率分析")
            text_1 = (f"报告期内，发行人各期的营业收入分别为{rev_t2:,.2f}万元、{rev_t1:,.2f}万元和{rev_t:,.2f}万元，"
                      f"营业成本分别为{cost_t2:,.2f}万元、{cost_t1:,.2f}万元和{cost_t:,.2f}万元，"
                      f"营业毛利率分别为{margins['T_2']:.2f}%、{margins['T_1']:.2f}%和{margins['T']:.2f}%。\n\n"
                      f"发行人以（）为主要业务，主要业务毛利水平较稳定。")
            st.markdown(text_1)
            st.code(text_1, language='text')

        with st.container(border=True):
            st.markdown("#### 📝 2、期间费用分析")
            text_2 = (f"报告期内，发行人期间费用总额分别为{period_expenses['T_2']:,.2f}万元、{period_expenses['T_1']:,.2f}万元和{period_expenses['T']:,.2f}万元，"
                      f"占发行人营业收入的比例分别为{pe_ratios['T_2']:.2f}%、{pe_ratios['T_1']:.2f}%和{pe_ratios['T']:.2f}%。\n\n"
                      f"报告期内，发行人期间费用主要为销售费用、管理费用、研发费用和财务费用，最近两年发行人期间费用较为稳定。\n\n")
            
            # 分项分析
            for name in exp_items:
                vals = get_val(name)
                # 占期间费用比例
                pct_pe_t = safe_pct(vals['T'], period_expenses['T'])
                pct_pe_t1 = safe_pct(vals['T_1'], period_expenses['T_1'])
                pct_pe_t2 = safe_pct(vals['T_2'], period_expenses['T_2'])
                # 占营收比例
                pct_rev_t = safe_pct(vals['T'], rev_t)
                pct_rev_t1 = safe_pct(vals['T_1'], rev_t1)
                pct_rev_t2 = safe_pct(vals['T_2'], rev_t2)
                
                text_2 += (f"报告期内，发行人发生{name}分别为{vals['T_2']:,.2f}万元、{vals['T_1']:,.2f}万元和{vals['T']:,.2f}万元，"
                           f"占期间费用的比例分别为{pct_pe_t2:.2f}%、{pct_pe_t1:.2f}%和{pct_pe_t:.2f}%，"
                           f"占营业收入的比重分别为{pct_rev_t2:.2f}%、{pct_rev_t1:.2f}%和{pct_rev_t:.2f}%。\n\n")
            
            st.markdown(text_2)
            st.code(text_2, language='text')

    with tab4:
        st.info("💡 **提示**：已自动生成关键盈利指标变动分析文案草稿。")
        # 1. 收入分析
        diff_rev_prev = rev_t1 - rev_t2
        diff_rev_curr = rev_t - rev_t1
        rev_text = (f"报告期各期，发行人营业收入分别为{rev_t2:,.2f}万元、{rev_t1:,.2f}万元和{rev_t:,.2f}万元。\n"
                    f"{d_t1}营业收入较{d_t2}变动{diff_rev_prev:,.2f}万元；\n"
                    f"{d_t}营业收入较{d_t1}变动{diff_rev_curr:,.2f}万元。\n"
                    f"变动主要原因为：（请结合业务规模、订单量、单价等因素分析）。")
        # 🟢 [修改]：Expander内改用 markdown + code
        with st.expander("📌 营业收入"): 
            st.markdown(rev_text)
            st.code(rev_text, language='text')
        
        # 2. 毛利率分析
        margin_text = (f"报告期各期，发行人毛利率分别为{margins['T_2']:.2f}%、{margins['T_1']:.2f}%、{margins['T']:.2f}%。\n"
                       f"发行人毛利率变动主要系：（请结合成本波动、产品定价策略等因素分析）。")
        with st.expander("📌 毛利率"): 
            st.markdown(margin_text)
            st.code(margin_text, language='text')

        # 3. 净利润分析
        net_t, net_t1, net_t2 = get_row_data(['净利润', '五、净利润'])
        net_text = (f"报告期各期，发行人净利润分别为{net_t2:,.2f}万元、{net_t1:,.2f}万元和{net_t:,.2f}万元。\n"
                    f"净利润变动趋势与利润总额变动趋势一致，变动原因主要为：（请补充非经常性损益或税务影响等原因）。")
        with st.expander("📌 净利润"): 
            st.markdown(net_text)
            st.code(net_text, language='text')


# ================= 5. 业务逻辑：财务指标分析 =================
def process_financial_ratios_tab(df_raw, word_data_list, d_labels):
    d_t, d_t1, d_t2 = d_labels
    
    # 🔥 核心修正：(显示名称, [搜索关键词], [排除关键词])
    metrics_config = [
        ("资产负债率（%）", ["资产负债率"], ["平均"]), # 排除“平均资产负债率”
        ("流动比率（倍）", ["流动比率"], None),
        ("速动比率（倍）", ["速动比率"], None),
        ("EBITDA（万元）", ["EBITDA", "息税折旧摊销前利润"], ["倍", "比", "率", "/", "%", "全部债务", "利息"]), # 排除比率类
        ("EBITDA利息保障倍数（倍）", ["EBITDA利息保障倍数", "利息保障倍数", "EBITDA利息倍数"], None)
    ]
    
    data_list = []
    data_map = {} 
    
    for display_name, search_kws, ex_kws in metrics_config:
        # 使用不带单位的关键词去模糊搜索
        row = find_row_fuzzy(df_raw, search_kws, exclude_keywords=ex_kws)
        
        val_t, val_t1, val_t2 = 0, 0, 0
        if row.name is not None:
            # 🔥 核心修正：应用智能单位转换
            is_ebitda = "EBITDA（万元）" in display_name
            is_ratio = "资产负债率" in display_name
            
            # 传入 subject_name 帮助判断单位
            val_t = smart_scale_convert(row['T'], row.name, is_ebitda, is_ratio)
            val_t1 = smart_scale_convert(row['T_1'], row.name, is_ebitda, is_ratio)
            val_t2 = smart_scale_convert(row['T_2'], row.name, is_ebitda, is_ratio)
            
            data_map[display_name] = {'T': val_t, 'T_1': val_t1, 'T_2': val_t2}
        
        if "EBITDA（万元）" in display_name:
            fmt_t = f"{val_t:,.2f}"
            fmt_t1 = f"{val_t1:,.2f}"
            fmt_t2 = f"{val_t2:,.2f}"
        else:
            fmt_t = f"{val_t:.2f}"
            fmt_t1 = f"{val_t1:.2f}"
            fmt_t2 = f"{val_t2:.2f}"
            
        data_list.append([display_name, fmt_t, fmt_t1, fmt_t2])

    df_display = pd.DataFrame(data_list, columns=["项目", d_t, d_t1, d_t2])
    df_display.set_index("项目", inplace=True)

    tab1, tab2, tab3 = st.tabs(["📋 指标数据", "📝 综述文案", "📝 变动分析文案"])

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

        # 🟢 [修改]：改用 container + markdown + code
        with st.container(border=True):
            st.markdown("#### 📝 偿债能力分析综述")
            
            text = f"1、资产负债率\n\n"
            text += f"报告期内，发行人的资产负债率分别为{alr['T_2']:.2f}%、{alr['T_1']:.2f}%和{alr['T']:.2f}%。\n\n"
            
            text += f"2、流动比率及速动比率\n\n"
            text += (f"报告期内，发行人的流动比率分别为{cr['T_2']:.2f}倍、{cr['T_1']:.2f}倍和{cr['T']:.2f}倍；"
                     f"报告期内，发行人的速动比率分别为{qr['T_2']:.2f}倍、{qr['T_1']:.2f}倍和{qr['T']:.2f}倍。\n\n")
            
            text += f"3、EBITDA利息保障倍数\n\n"
            text += (f"报告期内，发行人EBITDA分别为{ebitda['T_2']:,.2f}万元、{ebitda['T_1']:,.2f}万元和{ebitda['T']:,.2f}万元，"
                     f"发行人EBITDA利息保障倍数分别为{int_cov['T_2']:.2f}倍、{int_cov['T_1']:.2f}倍和{int_cov['T']:.2f}倍。")
            
            st.markdown(text)
            st.code(text, language='text')

    with tab3:
        st.info("💡 **提示**：已自动生成关键指标变动分析文案草稿。")
        prompts = [
            ("资产负债率", alr, "分析偿债风险变化"),
            ("流动比率", cr, "分析短期偿债能力"),
            ("EBITDA", ebitda, "分析盈利及获现能力")
        ]
        for name, data, task in prompts:
            # 根据趋势判断描述
            trend_text = ""
            if data['T'] > data['T_1']: trend_text = "有所上升"
            elif data['T'] < data['T_1']: trend_text = "有所下降"
            else: trend_text = "保持稳定"
            
            analysis_text = (f"报告期各期，发行人{name}分别为{data['T_2']:.2f}、{data['T_1']:.2f}和{data['T']:.2f}。\n"
                           f"报告期内，发行人{name}{trend_text}，主要系：（请结合资产负债结构或盈利能力分析）。")
            
            # 🟢 [修改]：Expander内改用 markdown + code
            with st.expander(f"📌 {name}"):
                st.markdown(analysis_text)
                st.code(analysis_text, language='text')

# ================= 3. 侧边栏 =================
with st.sidebar:
    st.title("🎛️ 操控台")
    analysis_page = st.radio(
        "请选择要生成的章节：", 
        ["(一) 资产结构分析", "(二) 负债结构分析", "(三) 现金流量分析", "(四) 财务指标分析", "(五) 盈利能力分析"]
    )
    st.markdown("---")
    
    uploaded_excel = st.file_uploader("Excel 底稿 (必须)", type=["xlsx", "xlsm"])
    
    # 💡 提示：Word 附注和高级设置已隐藏，系统将使用默认配置

# ================= 4. 主程序 =================

# --- ⚙️ 系统默认配置 (原高级设置内容) ---
# 由于删除了前端设置入口，此处定义默认值
DEFAULT_HEADER_ROW = 2  # 第3行
SHEET_CONFIG = {
    "asset": "1.合并资产表",
    "liab": "2.合并负债及权益表",
    "profit": "3.合并利润表",
    "cash": "4.合并现金流量表",
    "ratios": "5-3主要财务指标计算-方案3（专用公司债）"
}
# ------------------------------------

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
    2.  **自动分析**：上传即算，点击上方标签页切换 **数据表 / 文案 / 变动分析文案**。
    3.  **一键导出**：支持导出 **精排版 Word 表格** (宋体/加粗/1.5磅边框)。
    """)
    st.warning("👈 请先在左侧侧边栏上传 Excel 文件以开始使用。")

else:
    # ✅ 修复点 1：直接定义为空列表，不再尝试读取 uploaded_word_files
    word_data_list = [] 
    
    # ✅ 修复点 2：定义数据读取函数 (引用默认配置)
    def get_clean_data(target_sheet_name):
        try:
            # 使用默认的 HEADER_ROW = 2
            df, all_sheets_if_failed = fuzzy_load_excel(uploaded_excel, target_sheet_name, DEFAULT_HEADER_ROW)
            if df is None: return None, None, f"未找到 Sheet '{target_sheet_name}' (现有 Sheet: {all_sheets_if_failed})"
            
            # 尝试截取前几列 (假设格式标准)
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

    # --- 页面路由逻辑 ---

    if analysis_page == "(一) 资产结构分析":
        df_asset, d_labels, err = get_clean_data(SHEET_CONFIG["asset"])
        if df_asset is not None: process_analysis_tab(df_asset, word_data_list, "资产总计", "资产", d_labels)
        else: st.error(f"❌ 读取失败：{err}")

    elif analysis_page == "(二) 负债结构分析":
        df_liab, d_labels, err = get_clean_data(SHEET_CONFIG["liab"])
        if df_liab is not None:
            total_name = "负债合计" 
            if not df_liab.index.str.contains(total_name).any(): total_name = "负债总计"
            process_analysis_tab(df_liab, word_data_list, total_name, "负债", d_labels)
        else: st.error(f"❌ 读取失败：{err}")

    elif analysis_page == "(三) 现金流量分析":
        df_cash, d_labels, err = get_clean_data(SHEET_CONFIG["cash"])
        if df_cash is not None:
            process_cash_flow_tab(df_cash, word_data_list, d_labels)
        else: st.error(f"❌ 读取失败：{err}")

    elif analysis_page == "(四) 财务指标分析":
        # 财务指标表通常表头不固定，使用 fuzzy_load_excel 的内部逻辑
        df_ratios, d_labels = fuzzy_load_excel(uploaded_excel, SHEET_CONFIG["ratios"], DEFAULT_HEADER_ROW)
        if df_ratios is not None:
            process_financial_ratios_tab(df_ratios, word_data_list, d_labels)
        else: 
            st.error(f"❌ 读取失败：未找到 Sheet '{SHEET_CONFIG['ratios']}'")

    elif analysis_page == "(五) 盈利能力分析":
        df_profit, d_labels, err = get_clean_data(SHEET_CONFIG["profit"])
        if df_profit is not None:
            process_profitability_tab(df_profit, word_data_list, d_labels)
        else: st.error(f"❌ 读取失败：{err}")
