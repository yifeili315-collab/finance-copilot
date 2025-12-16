import streamlit as st
import pandas as pd
import re
from docx import Document
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
2. 上传 **Word 附注**（可选，用于生成原因分析）。
3. 系统会自动计算数据，并生成 **AI 提问指令**。
4. 点击指令右上角的 **📄 复制按钮**，发送给你常用的 AI (ChatGPT/DeepSeek) 即可。
""")

# ================= 2. 侧边栏：文件上传 =================
with st.sidebar:
    st.header("📂 请上传文件")
    
    # 1. Excel (必须)
    uploaded_excel = st.file_uploader("1. 上传 Excel 底稿 (必须)", type=["xlsx", "xlsm"])
    
    # 2. Word (可选)
    uploaded_word = st.file_uploader("2. 上传 Word 附注 (可选)", type=["docx"], help="如果上传，分析将包含原因；如果不传，只生成数据描述。")
    
    st.info("💡 提示：数据只在浏览器本地处理，不会上传给第三方 AI，绝对安全。")
    
    # 允许用户调整表头行
    header_row = st.number_input("Excel表头所在行 (默认2，即第3行)", value=2, min_value=0)

# ================= 3. 核心逻辑函数 =================

def load_word_context(file_obj):
    """读取 Word 文件流"""
    if file_obj is None: return ""
    try:
        doc = Document(file_obj)
        full_text = []
        for para in doc.paragraphs:
            clean = para.text.strip()
            if len(clean) > 5:
                full_text.append(clean)
        return "\n".join(full_text)
    except Exception as e:
        st.error(f"Word 读取失败: {e}")
        return ""

def find_context(subject, full_text):
    """RAG 检索"""
    if not full_text: return ""
    clean_sub = subject.replace(" ", "")
    idx = full_text.find(clean_sub)
    if idx == -1: return "（附注中未检索到该科目名称）"
    start = max(0, idx - 600)
    end = min(len(full_text), idx + 1000)
    return full_text[start:end].replace('\n', ' ')

def clean_date_label(header_str):
    """清洗日期标签"""
    s = str(header_str).replace('\n', '')
    year = re.search(r'(\d{4})', s)
    y_str = year.group(1) if year else "T"
    suffix = "6月末" if ("一期" in s or "6月" in s) else "年末"
    return f"{y_str}年{suffix}"

# 安全计算占比函数
def safe_pct(num, denom):
    return (num / denom * 100) if denom != 0 else 0.0

# ================= 4. 主程序逻辑 =================

if uploaded_excel:
    # 🌟 修改点 1：文案优化 - 明确告知用户分析已完成
    st.success("✅ 分析完成！已生成数据综述与 AI 指令，请查看下方结果。")
    
    try:
        # --- 1. 读取并处理 Excel ---
        df = pd.read_excel(uploaded_excel, sheet_name='1.合并资产表', header=header_row)
        df = df.iloc[:, [0, 4, 5, 6]]
        
        orig_cols = df.columns.tolist()
        d_t = clean_date_label(orig_cols[1])
        d_t1 = clean_date_label(orig_cols[2])
        d_t2 = clean_date_label(orig_cols[3])
        
        df.columns = ['科目', 'T', 'T_1', 'T_2']
        
        # 数据清洗
        df = df.dropna(subset=['科目'])
        df['科目'] = df['科目'].astype(str).str.strip()
        for c in ['T', 'T_1', 'T_2']:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
            
        df.set_index('科目', inplace=True)
        
        # 提取关键行
        try:
            total_assets = df[df.index.str.contains('资产总计|资产总额')].iloc[0]
            curr_assets = df[df.index.str.contains('流动资产合计')].iloc[0]
            non_curr_assets = df[df.index.str.contains('非流动资产合计')].iloc[0]
        except IndexError:
            st.error("❌ 未找到 '资产总计' 等关键行，请检查 Excel 科目名称。")
            st.stop()
            
        # 计算 T 期占比
        if total_assets['T'] != 0:
            df['占比_T'] = df['T'] / total_assets['T']
        else:
            df['占比_T'] = 0.0
        
        # --- 2. 尝试读取 Word ---
        word_text = ""
        has_word = False
        if uploaded_word:
            word_text = load_word_context(uploaded_word)
            if word_text:
                has_word = True
            else:
                st.warning("Word 文件为空或读取失败，将生成纯数据指令。")

        # ================= 5. 结果展示 =================
        
        tab1, tab2, tab3 = st.tabs(["📋 1. 资产明细", "📝 2. 综述文案", "🤖 3. AI 写作指令"])
        
        # --- Tab 1: 明细表 ---
        with tab1:
            st.markdown("### 资产结构明细")
            display_df = df.copy()
            display_df['占比_T(%)'] = (display_df['占比_T'] * 100).apply(lambda x: f"{x:.2f}%")
            st.dataframe(
                display_df[['T', 'T_1', 'T_2', '占比_T(%)']].style.format(
                    subset=['T', 'T_1', 'T_2'], 
                    formatter="{:,.2f}"
                )
            )

        # --- Tab 2: 综述文案 ---
        with tab2:
            st.subheader("资产结构总体分析")
            st.markdown("👇 **直接复制到报告：**")
            
            exclude = ['合计', '总计', '总额']
            detail_df = df[~df.index.str.contains('|'.join(exclude))]
            top_5 = detail_df.sort_values(by='T', ascending=False).head(5).index.tolist()
            top_5_str = "、".join(top_5)
            
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
                f"在总资产构成中，公司资产主要为 **{top_5_str}** 等。"
            )
            st.code(text_overview, language='text')

        # --- Tab 3: AI 指令 (核心优化点) ---
        with tab3:
            st.subheader("🤖 重点科目分析指令 (Copilot 模式)")
            st.caption("👉 点击代码块右上角的 **📄 复制**，粘贴给 DeepSeek 或 ChatGPT。")
            
            if not has_word:
                st.info("ℹ️ 未上传 Word，生成【纯数据分析指令】。")
            
            major_subjects = detail_df[detail_df['占比_T'] > 0.01].index.tolist()
            
            for subject in major_subjects:
                row = df.loc[subject]
                v_t2, v_t1, v_t = row['T_2'], row['T_1'], row['T']
                r_t2 = safe_pct(v_t2, total_assets['T_2'])
                r_t1 = safe_pct(v_t1, total_assets['T_1'])
                r_t = safe_pct(v_t, total_assets['T'])
                
                diff = v_t - v_t1
                pct = safe_pct(diff, v_t1)
                
                # 🌟 修改点 2：智能判断增幅还是降幅
                if diff >= 0:
                    direction = "增加"
                    pct_label = "增幅"
                else:
                    direction = "减少"
                    pct_label = "降幅"
                
                # 基础数据 (Prompt) - 这里用了 pct_label 变量
                prompt_base = f"""【任务】：请分析“{subject}”的变动情况。

【1. 财务具体科目数据 (Trend)】
{d_t2}、{d_t1}及{d_t}，发行人{subject}余额分别为{v_t2:,.2f}万元、{v_t1:,.2f}万元和{v_t:,.2f}万元，占总资产的比例分别为{r_t2:.2f}%、{r_t1:.2f}%和{r_t:.2f}%。

【2. 财务硬数据变动 (Analysis)】
截至{d_t}，发行人{subject}较{d_t1}{direction}{abs(diff):,.2f}万元，{pct_label}为{abs(pct):.2f}%。"""

                # 智能组合
                if has_word:
                    context = find_context(subject, word_text)
                    prompt_final = prompt_base + f"""

【3. Word 附注软信息 (Context)】
{context}

【4. 写作指令】
1. 请将上述数据整理成一段通顺的财务分析文字。
2. 紧接着，请根据第3部分分析变动原因（即“主要系...所致”）。
3. 如果Part 3没提到具体原因，请直接写“主要系业务规模变动所致”，严禁瞎编。"""
                else:
                    prompt_final = prompt_base + "\n\n【指令】请将上述数据整理成一段通顺的财务分析文字。"

                with st.expander(f"📌 {subject} (占比 {r_t:.2f}%)"):
                    st.code(prompt_final, language='text')

    except Exception as e:
        st.error(f"解析出错: {e}")
        st.info("请检查 Excel 格式是否与模版一致。")

else:
    st.info("👈 请在左侧上传文件以开始分析")
