# -*- coding: utf-8 -*-
"""
BERTScore Evaluation Script
อ่าน summaries จาก 75clip_mapper.xlsx แล้วคำนวณ BERTScore
เปรียบเทียบ temp0.0, temp0.2, temp0.4, temp0.6 กับ ref_ChatGPT, ref_Gemini
"""

import pandas as pd
from bert_score import score
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# ================== CONFIG ==================
INPUT_FILE = r"d:\final_project_\analysis\75clip_mapper.xlsx"
OUTPUT_FILE = r"d:\final_project_\analysis\evaluation_results.xlsx"

TEMP_COLUMNS = ['temp0.0', 'temp0.2', 'temp0.4', 'temp0.6']
REF_COLUMNS = ['ref_ChatGPT', 'ref_Gemini']

# ================== MAIN ==================
def main():
    print("=" * 60)
    print("BERTScore Evaluation Script")
    print("=" * 60)
    
    # 1. Load data
    print("\n📖 Loading data from Excel...")
    df = pd.read_excel(INPUT_FILE)
    print(f"   Loaded {len(df)} rows")
    print(f"   Columns: {df.columns.tolist()}")
    
    # 2. Calculate BERTScore for each combination
    print(f"\n{'='*60}")
    print("Calculating BERTScore...")
    print(f"{'='*60}")
    
    all_results = []
    
    for idx, row in df.iterrows():
        video_url = row.get('YoutubeUrl', '')
        category = row.get('Category', '')
        duration = row.get('Duration(min)', '')
        
        # Extract video ID from URL
        video_id = ""
        if "v=" in str(video_url):
            video_id = str(video_url).split("v=")[1].split("&")[0]
        elif "youtu.be/" in str(video_url):
            video_id = str(video_url).split("youtu.be/")[1].split("?")[0]
        elif "shorts/" in str(video_url):
            video_id = str(video_url).split("shorts/")[1].split("?")[0]
        
        print(f"\n[{idx+1}/{len(df)}] Processing: {video_id[:20]}... ({category})")
        
        for ref_col in REF_COLUMNS:
            ref_summary = str(row.get(ref_col, '')).strip()
            ref_tool = ref_col.replace('ref_', '')
            
            if not ref_summary or ref_summary == 'nan' or len(ref_summary) < 10:
                print(f"   ⚠️ Skipping {ref_tool} - no reference summary")
                continue
            
            for temp_col in TEMP_COLUMNS:
                cand_summary = str(row.get(temp_col, '')).strip()
                temp_value = temp_col.replace('temp', '')
                
                if not cand_summary or cand_summary == 'nan' or len(cand_summary) < 10:
                    print(f"   ⚠️ Skipping {temp_col} - no candidate summary")
                    continue
                
                try:
                    # Calculate BERTScore
                    P, R, F1 = score(
                        [cand_summary],
                        [ref_summary],
                        lang="th",
                        verbose=False,
                        rescale_with_baseline=True
                    )
                    
                    result = {
                        "video_id": video_id,
                        "category": category,
                        "duration_min": duration,
                        "whisper_temp": float(temp_value),
                        "reference_tool": ref_tool,
                        "precision": round(P.mean().item(), 4),
                        "recall": round(R.mean().item(), 4),
                        "f1": round(F1.mean().item(), 4),
                        "cand_length": len(cand_summary),
                        "ref_length": len(ref_summary),
                    }
                    all_results.append(result)
                    
                    print(f"   temp{temp_value} vs {ref_tool}: F1={result['f1']:.4f}")
                    
                except Exception as e:
                    print(f"   ❌ Error: {e}")
    
    # 3. Create DataFrames
    print(f"\n{'='*60}")
    print("Creating result sheets...")
    print(f"{'='*60}")
    
    df_results = pd.DataFrame(all_results)
    
    if df_results.empty:
        print("❌ No results to save!")
        return
    
    # 4. Save to Excel with multiple sheets
    print(f"\n💾 Saving to {OUTPUT_FILE}...")
    
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        # Sheet 1: All BERTScore Results
        df_results.to_excel(writer, sheet_name='BERTScore_All', index=False)
        print(f"   ✅ Sheet 'BERTScore_All': {len(df_results)} rows")
        
        # Sheet 2: Pivot by Video (F1 scores)
        df_pivot = df_results.pivot_table(
            values='f1',
            index=['video_id', 'category'],
            columns=['whisper_temp', 'reference_tool'],
            aggfunc='mean'
        ).reset_index()
        df_pivot.to_excel(writer, sheet_name='Pivot_by_Video', index=False)
        print(f"   ✅ Sheet 'Pivot_by_Video': F1 by video and temp")
        
        # Sheet 3: Average by Temp
        df_avg_temp = df_results.groupby(['whisper_temp', 'reference_tool']).agg({
            'precision': 'mean',
            'recall': 'mean',
            'f1': 'mean'
        }).reset_index()
        df_avg_temp.columns = ['whisper_temp', 'reference_tool', 'avg_precision', 'avg_recall', 'avg_f1']
        df_avg_temp = df_avg_temp.round(4)
        df_avg_temp.to_excel(writer, sheet_name='Average_by_Temp', index=False)
        print(f"   ✅ Sheet 'Average_by_Temp': Average scores by temperature")
        
        # Sheet 4: Average by Category
        df_avg_cat = df_results.groupby(['category', 'whisper_temp']).agg({
            'f1': 'mean'
        }).reset_index()
        df_avg_cat.columns = ['category', 'whisper_temp', 'avg_f1']
        df_avg_cat = df_avg_cat.round(4)
        df_avg_cat_pivot = df_avg_cat.pivot(index='category', columns='whisper_temp', values='avg_f1').reset_index()
        df_avg_cat_pivot.to_excel(writer, sheet_name='Average_by_Category', index=False)
        print(f"   ✅ Sheet 'Average_by_Category': Average F1 by category and temp")
        
        # Sheet 5: Summary Statistics
        summary_stats = []
        for temp in [0.0, 0.2, 0.4, 0.6]:
            temp_data = df_results[df_results['whisper_temp'] == temp]
            if not temp_data.empty:
                summary_stats.append({
                    'whisper_temp': temp,
                    'count': len(temp_data),
                    'f1_mean': temp_data['f1'].mean(),
                    'f1_std': temp_data['f1'].std(),
                    'f1_min': temp_data['f1'].min(),
                    'f1_max': temp_data['f1'].max(),
                    'precision_mean': temp_data['precision'].mean(),
                    'recall_mean': temp_data['recall'].mean(),
                })
        df_summary = pd.DataFrame(summary_stats).round(4)
        df_summary.to_excel(writer, sheet_name='Summary_Stats', index=False)
        print(f"   ✅ Sheet 'Summary_Stats': Overall statistics")
    
    print(f"\n{'='*60}")
    print(f"✅ Results saved to: {OUTPUT_FILE}")
    print(f"{'='*60}")
    
    # Print quick summary
    print("\n📊 Quick Summary (Average F1 by Temp):")
    print(df_avg_temp.to_string(index=False))

if __name__ == "__main__":
    main()
