import streamlit as st
import pandas as pd
from collections import defaultdict
import sys
import re
import io
from datetime import datetime
import tempfile
import os

# Streamlit App Title and Description
st.title("品牌广告 Header 生成工具")
st.markdown("""
### 代码内容说明
此工具基于提供的 Python 脚本 `test 品牌所有广告 集合版.py` 开发，用于从上传的 Excel 文件（默认 sheet: '品牌广告'）中提取全局设置、活动数据和关键词信息，生成广告 Header 文件。  
**主要功能：**  
- 支持（品牌旗舰店、商品集、商品详情页）主题的动态区域检测和数据提取。  
- 处理广告活动、广告组、视频/商品集广告、关键词、否定关键词、商品定向等行生成。  
- 自动填充默认值（如预算类型 '每日'、状态 '已启用'）。  
- 检测重复否定关键词并暂停生成（打印警告）。  
- 输出 27 列标准 Header 格式的 Excel 文件。  

**使用步骤：**  
1. 上传 Excel 文件（文件名任意，需包含 '品牌广告' sheet）。  
2. 点击 "生成 Header 文件" 按钮。  
3. 下载生成的 "header-品牌 YYYY-MM-DD HH:MM.xlsx" 文件。  

**注意：**  
- 文件需符合脚本预期结构（A 列主题行、B 列活动名称等）。  
- 如遇错误（如未找到主题），页面将显示日志。  
- 生成时间精确到分钟（基于当前时间）。  
""")

# File Uploader
uploaded_file = st.file_uploader("上传 Excel 文件", type=['xlsx', 'xls'])

# Function from the original script (copied and adapted)
def generate_header_for_sbv_brand_store(uploaded_bytes, sheet_name='品牌广告'):
    # Create a temporary file from bytes
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        tmp.write(uploaded_bytes)
        input_file = tmp.name
    
    try:
        # Read the entire file, header=0
        df_survey = pd.read_excel(input_file, sheet_name=sheet_name, header=0)
        st.write(f"成功读取文件，数据形状：{df_survey.shape}")
        st.write(f"列名列表: {list(df_survey.columns)}")
    except FileNotFoundError:
        st.error(f"错误：未找到文件。请确保文件包含 '{sheet_name}' sheet。")
        os.unlink(input_file)
        return None
    except Exception as e:
        st.error(f"读取文件时出错：{e}")
        os.unlink(input_file)
        return None
    
    # Fill NaN with empty string
    df_survey = df_survey.fillna('')
    
    # 修复1: 移除区域检测（你的Survey扁平，无主题行）
    # global_limit = ...  # 注释掉，用全表: global_limit = len(df_survey)

    # 修复2: 全局设置从数据行1起（iloc[1:, 0:2]）
    global_settings = {}
    data_start = 1  # Skip header row0
    for i in range(data_start, min(data_start + 20, len(df_survey))):
        label = str(df_survey.iloc[i, 0]).strip()
        value = str(df_survey.iloc[i, 1]).strip() if len(df_survey.columns) > 1 else ''
        if '品牌实体编号' in label or 'ENTITY' in label.upper():
            global_settings['entity_id'] = value
        elif '品牌名称' in label:
            global_settings['brand_name'] = value
        elif '预算类型' in label:
            global_settings['budget_type'] = value if value else '每日'
        elif '创意素材标题' in label:
            global_settings['creative_title'] = value
        elif '落地页 URL' in label:
            global_settings['landing_url'] = value
    st.write(f"全局设置: {global_settings}")
    
    # 修复3: header_row_full = df_survey.columns.tolist()  # 正确用列名
    header_row_full = df_survey.columns.tolist()
    keyword_columns = [col for col in header_row_full if isinstance(col, str) and ('精准词' in col or '广泛词' in col or '否' in col)]
    st.write(f"关键词相关列: {keyword_columns}")
    
    # Identify keyword categories like in test SB.py
    keyword_categories = set()
    for col in keyword_columns:
        col_lower = str(col).lower()
        if '/' in col_lower:
            parts = col_lower.split('/')
            if len(parts) > 0 and parts[0]:
                keyword_categories.add(parts[0].strip())
            if len(parts) > 1 and parts[1]:
                chinese_part = parts[1].split('-')[0].strip() if '-' in parts[1] else parts[1].strip()
                keyword_categories.add(chinese_part)
        else:
            for suffix in ['精准词', '广泛词', '精准', '广泛']:
                if col_lower.endswith(suffix):
                    prefix = col_lower[:-len(suffix)].strip()
                    if prefix:
                        keyword_categories.add(prefix)
                        break
    keyword_categories.update(['suzhu', '宿主', 'host', 'case', '包', '对手', 'tape'])
    st.write(f"识别到的关键词类别: {keyword_categories}")
    
    # Negative keywords extraction: map to specific columns like test SB.py
    # Col indices mapping
    col_indices = {
        'W': df_survey.columns.get_loc('宿主精准-否精准') if '宿主精准-否精准' in df_survey.columns else None,
        'X': df_survey.columns.get_loc('宿主精准-否词组') if '宿主精准-否词组' in df_survey.columns else None,
        'AA': df_survey.columns.get_loc('宿主广泛-否精准') if '宿主广泛-否精准' in df_survey.columns else None,
        'AB': df_survey.columns.get_loc('宿主广泛-否词组') if '宿主广泛-否词组' in df_survey.columns else None,
        'Y': df_survey.columns.get_loc('case精准-否精准') if 'case精准-否精准' in df_survey.columns else None,
        'Z': df_survey.columns.get_loc('case精准-否词组') if 'case精准-否词组' in df_survey.columns else None,
        'AC': df_survey.columns.get_loc('case广泛-否精准') if 'case广泛-否精准' in df_survey.columns else None,
        'AD': df_survey.columns.get_loc('case广泛-否词组') if 'case广泛-否词组' in df_survey.columns else None,
    }
    
    # Col names for logging
    col_names_dict = {
        'W': '宿主精准-否精准',
        'X': '宿主精准-否词组',
        'AA': '宿主广泛-否精准',
        'AB': '宿主广泛-否词组',
        'Y': 'case精准-否精准',
        'Z': 'case精准-否词组',
        'AC': 'case广泛-否精准',
        'AD': 'case广泛-否词组'
    }
    
    # Extract neg_asin and neg_brand from specific columns
    neg_asin = []
    neg_brand = []
    neg_asin_col = None
    neg_brand_col = None
    for col_idx, col_name in enumerate(df_survey.columns):
        if '否定asin' in str(col_name).lower():
            neg_asin_col = col_idx
        elif '否品牌' in str(col_name).lower():
            neg_brand_col = col_idx
    if neg_asin_col is not None:
        neg_asin = [str(x).strip() for x in df_survey.iloc[:, neg_asin_col].dropna() if str(x).strip()]
        neg_asin = list(dict.fromkeys(neg_asin))
    if neg_brand_col is not None:
        neg_brand = [str(int(x)).strip() for x in df_survey.iloc[:, neg_brand_col].dropna() if str(x).strip()]
        neg_brand = list(dict.fromkeys(neg_brand))
    st.write(f"否定ASIN: {neg_asin}")
    st.write(f"否品牌: {neg_brand}")
    
    # Output columns from original test.py (27 columns)
    output_columns = [
        '产品', '实体层级', '操作', '广告活动编号', '广告组编号', '广告编号', 
        '广告活动名称', '广告组名称', '广告名称', '状态', '品牌实体编号', 
        '预算类型', '预算', '商品位置', '竞价', '关键词文本', '匹配类型', '拓展商品投放编号', 
        '落地页 URL', '落地页类型', '品牌名称', '同意翻译', '品牌徽标素材编号', 
        '创意素材标题', '创意素材 ASIN', '视频素材编号', '自定义图片'
    ]
    
    product = '品牌推广'
    operation = 'Create'
    status = '已启用'
    
    rows = []
    
    default_bid = 0.6
    
    # 修复4: 活动循环用全表 (移除global_limit限)
    for col_idx, col in enumerate(header_row_full):
        if isinstance(col, str) and 'sb_' in col.lower() and '25/11/18' in col:  # 你的campaign模式
            campaign_name = str(col).strip()
            is_exact = 'exact' in campaign_name.lower()
            is_broad = 'broad' in campaign_name.lower()
            is_asin = 'asin' in campaign_name.lower()
            if not any([is_exact, is_broad, is_asin]): continue  # Skip invalid

            # matched_category (不变)
            matched_category = None
            for cat in ['suzhu', '宿主', 'host', 'case', '包']:
                if cat.lower() in campaign_name.lower():
                    matched_category = cat.lower()
                    break
            st.write(f"活动: {campaign_name}, category: {matched_category}, type: {'exact' if is_exact else 'broad' if is_broad else 'asin'}")

            # 广告活动/组/商品集行 (不变，添加)
            row_campaign = [product, '广告活动', operation, campaign_name, '', campaign_name, '', '', status, global_settings.get('entity_id', ''), 
                            global_settings.get('budget_type', '每日'), '10', '在亚马逊上出售', '', '', '', '', '', global_settings.get('landing_url', ''), '品牌旗舰店', 
                            global_settings.get('brand_name', ''), 'False', '', global_settings.get('creative_title', ''), '', '', '']
            rows.append(row_campaign)

            row_group = [product, '广告组', operation, campaign_name, campaign_name, '', campaign_name, '', status, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
            rows.append(row_group)

            row_product = [product, '商品集广告', operation, campaign_name, campaign_name, campaign_name, campaign_name, '', status, '', '', '', '', '', '', '', '', global_settings.get('landing_url', ''), 
                           '品牌旗舰店', global_settings.get('brand_name', ''), 'False', '', global_settings.get('creative_title', ''), '', '', '']
            rows.append(row_product)

            cpc = default_bid  # 或从设置
            match_type = '精准' if is_exact else '广泛'

            # 修复5: 关键词提取 — 用正确列loc，从全数据 (iloc[1:]) 提取非空
            keywords = []
            if matched_category in ['suzhu', '宿主', 'host']:
                precise_col = 'suzhu/宿主/host-精准词'
                broad_col = 'suzhu/宿主/host-广泛词'
            elif matched_category in ['case', '包']:
                precise_col = 'case/包-精准词'
                broad_col = 'case/包-广泛词'
            else:
                continue

            # Extract from data rows (skip header)
            data_slice = df_survey.iloc[1:]
            if precise_col in header_row_full:
                keywords += data_slice[precise_col].dropna().str.strip().str.cat(sep='\n').split('\n')  # All non-empty
                keywords = [kw for kw in keywords if kw]
            if is_broad and broad_col in header_row_full:
                broad_kws = data_slice[broad_col].dropna().str.strip().str.cat(sep='\n').split('\n')
                keywords += [kw for kw in broad_kws if kw]
            keywords = list(dict.fromkeys(keywords))  # Dedup
            st.write(f"  提取关键词: {keywords[:3]} (总{len(keywords)})")

            if keywords:
                for kw in keywords:
                    row_keyword = [product, '关键词', operation, campaign_name, campaign_name, '', '', '', status, '', '', '', '', cpc, kw, match_type, '', '', '', '', '', '', '', '', '', '']
                    rows.append(row_keyword)
            else:
                # Only fallback if truly empty
                row_keyword = [product, '关键词', operation, campaign_name, campaign_name, '', '', '', status, '', '', '', '', cpc, 'default_kw', match_type, '', '', '', '', '', '', '', '', '', '']
                rows.append(row_keyword)
                st.warning(f"  无关键词 for {campaign_name}，用default_kw")
            
            # Negative keywords: dynamic like test SB.py, with specific column selection
            if matched_category:
                # Select columns based on category and type
                selected_cols = []
                if matched_category in ['suzhu', '宿主', 'host']:
                    if is_exact:
                        selected_cols = ['W', 'X']
                    elif is_broad:
                        selected_cols = ['AA', 'AB']
                elif matched_category in ['case', '包']:
                    if is_exact:
                        selected_cols = ['Y', 'Z']
                    elif is_broad:
                        selected_cols = ['AC', 'AD']
                
                # Collect data, track sources for duplicates
                neg_data_sources = {
                    '否定精准匹配': defaultdict(list),  # kw -> [col_keys]
                    '否定词组': defaultdict(list)
                }
                for col_key in selected_cols:
                    if col_indices.get(col_key) is not None:
                        col_idx = col_indices[col_key]
                        col_data = [str(kw).strip() for kw in df_survey.iloc[:, col_idx].dropna() if str(kw).strip()]
                        col_data = list(dict.fromkeys(col_data))  # column dedup
                        m_type = '否定精准匹配' if col_key in ['W', 'AA', 'Y', 'AC'] else '否定词组'
                        for kw in col_data:
                            neg_data_sources[m_type][kw].append(col_key)
                
                # Check duplicates: kw with multiple sources
                duplicates_detected = False
                for m_type, kw_sources in neg_data_sources.items():
                    for kw, sources in kw_sources.items():
                        if len(sources) > 1:
                            duplicates_detected = True
                            source_names = [col_names_dict.get(s, s) for s in sources]
                            st.error(f"\n=== 检测到重复否定关键词 ===")
                            st.error(f"活动: {campaign_name}")
                            st.error(f"类型: {m_type}")
                            st.error(f"重复关键词: '{kw}'")
                            st.error(f"来源列: {', '.join(source_names)}")
                            st.error(f"原因: 该关键词在多个否定列中出现，导致生成重复行。请检查 survey 文件的这些列并清理重复值。")
                            st.error("暂停生成 header 表。")
                            os.unlink(input_file)
                            return None  # Pause generation
                
                st.write("\n=== 重复检测完成（无重复）===")
                
                # Generate rows: deduped kws
                for m_type, kw_sources in neg_data_sources.items():
                    kws = list(kw_sources.keys())
                    if kws:
                        st.write(f"  {m_type} 否定关键词数量: {len(kws)}")
                    for kw in kws:
                        row_neg = [product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '', status, 
                                   '', '', '', '', '', kw, m_type, '', '', '', '', '', '', '', '', '', '']
                        rows.append(row_neg)

            # ASIN group: generate 商品定向 and 否定商品定向
            if is_asin:
                # 商品定向: exact column match to campaign_name
                asin_targets = []
                for col in df_survey.columns:
                    if str(col).strip() == str(campaign_name):
                        col_idx = df_survey.columns.get_loc(col)
                        if col_idx is not None:
                            asin_targets = [str(asin).strip() for asin in df_survey.iloc[:, col_idx].dropna() if str(asin).strip()]
                            asin_targets = list(dict.fromkeys(asin_targets))
                            st.write(f"  商品定向 ASIN 数量: {len(asin_targets)} (示例: {asin_targets[:2] if asin_targets else '无'})")
                            break
                
                if asin_targets:
                    for asin in asin_targets:
                        row_product_target = [product, '商品定向', operation, campaign_name, campaign_name, '', '', campaign_name, '', status, 
                                              '', '', '', '', cpc, '', '', f'asin="{asin}"', '', '', '', '', '', '', '', '', '']
                        rows.append(row_product_target)
                
                # 否定商品定向: from global neg_asin and neg_brand
                for neg in neg_asin:
                    row_neg_product = [product, '否定商品定向', operation, campaign_name, campaign_name, '', '', campaign_name, '', status, 
                                       '', '', '', '', '', '', '', f'asin="{neg}"', '', '', '', '', '', '', '', '', '']
                    rows.append(row_neg_product)
                
                for negb in neg_brand:
                    row_neg_brand = [product, '否定商品定向', operation, campaign_name, campaign_name, '', '', campaign_name, '', status, 
                                     '', '', '', '', '', '', '', f'brand="{negb}"', '', '', '', '', '', '', '', '', '']
                    rows.append(row_neg_brand)
    
    # Create DF
    df_header = pd.DataFrame(rows, columns=output_columns)
    df_header = df_header.fillna('')
    
    # Save to BytesIO for download
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        df_header.to_excel(writer, index=False, sheet_name='Sheet1')
    output_buffer.seek(0)
    
    st.success(f"生成完成！总行数：{len(rows)}")
    
    # Cleanup temp file
    os.unlink(input_file)
    
    return output_buffer

# Generate Button
if uploaded_file is not None:
    if st.button("生成 Header 文件"):
        with st.spinner("正在处理文件..."):
            output_buffer = generate_header_for_sbv_brand_store(uploaded_file.read())
            if output_buffer is not None:
                # Generate filename with current time (precise to minute)
                now = datetime.now()
                timestamp = now.strftime("%Y-%m-%d %H:%M")
                filename = f"header-品牌 {timestamp}.xlsx"
                
                st.download_button(
                    label="下载生成的 Header 文件",
                    data=output_buffer.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
else:
    st.info("请上传 Excel 文件以开始生成。")
