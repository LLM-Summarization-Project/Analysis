# -*- coding: utf-8 -*-
"""
Extract summaries from stats_data.csv and add them to 75clip_mapper.xlsx
Each YouTube URL in mapper has 4 corresponding rows in stats_data (one per whisperTemp: 0.0, 0.2, 0.4, 0.6)
Creates new columns: temp0.0, temp0.2, temp0.4, temp0.6 with summary content
"""

import pandas as pd
import os
import re

# Paths
MAPPER_FILE = r"d:\final_project_\analysis\75clip_mapper.xlsx"
STATS_FILE = r"d:\final_project_\analysis\stats_data.csv"
OUTPUT_FILE = r"d:\final_project_\analysis\75clip_mapper.xlsx"  # Update original file

# Base path for summary files (Docker path -> Windows path mapping)
DOCKER_BASE = "/app/outputs"
WINDOWS_BASE = r"d:\final_project_\summarize-backend\outputs"

def normalize_youtube_url(url: str) -> str:
    """Normalize YouTube URL to extract video ID for matching"""
    if not url or pd.isna(url):
        return ""
    url = str(url).strip()
    
    # Extract video ID from various YouTube URL formats
    patterns = [
        r'(?:v=|/v/|youtu\.be/|/embed/|/shorts/)([a-zA-Z0-9_-]{11})',
        r'^([a-zA-Z0-9_-]{11})$'
    ]
    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)
    return url

def docker_path_to_windows(docker_path: str) -> str:
    """Convert Docker path to Windows path"""
    if not docker_path or pd.isna(docker_path):
        return ""
    return docker_path.replace(DOCKER_BASE, WINDOWS_BASE).replace("/", "\\")

def read_summary_file(summary_path: str) -> str:
    """Read summary content from file"""
    if not summary_path:
        return ""
    
    windows_path = docker_path_to_windows(summary_path)
    
    try:
        if os.path.exists(windows_path):
            with open(windows_path, 'r', encoding='utf-8') as f:
                return f.read().strip()
        else:
            print(f"⚠️ File not found: {windows_path}")
            return ""
    except Exception as e:
        print(f"❌ Error reading {windows_path}: {e}")
        return ""

def main():
    print("📖 Loading mapper file...")
    mapper_df = pd.read_excel(MAPPER_FILE)
    print(f"   Loaded {len(mapper_df)} rows from mapper")
    
    print("📖 Loading stats_data.csv...")
    stats_df = pd.read_csv(STATS_FILE)
    print(f"   Loaded {len(stats_df)} rows from stats_data")
    
    # Add normalized video ID for matching
    mapper_df['video_id'] = mapper_df['YoutubeUrl'].apply(normalize_youtube_url)
    stats_df['video_id'] = stats_df['youtubeUrl'].apply(normalize_youtube_url)
    
    # Check whisperTemp values available
    temps = stats_df['whisperTemp'].dropna().unique()
    print(f"   Available whisperTemp values: {sorted(temps)}")
    
    # Initialize new columns
    temp_columns = ['temp0.0', 'temp0.2', 'temp0.4', 'temp0.6']
    for col in temp_columns:
        mapper_df[col] = ""
    
    # Process each row in mapper
    print("\n🔄 Processing summaries...")
    found_count = 0
    missing_count = 0
    
    for idx, row in mapper_df.iterrows():
        video_id = row['video_id']
        if not video_id:
            continue
        
        # Find matching rows in stats_data for this video
        matching = stats_df[stats_df['video_id'] == video_id]
        
        if matching.empty:
            missing_count += 1
            print(f"⚠️ No match for: {row['YoutubeUrl'][:50]}...")
            continue
        
        found_count += 1
        
        # Get summary for each temperature
        for temp in [0.0, 0.2, 0.4, 0.6]:
            temp_rows = matching[matching['whisperTemp'] == temp]
            
            if not temp_rows.empty:
                # Get the first matching row's summary
                summary_path = temp_rows.iloc[0]['summaryPath']
                summary_content = read_summary_file(summary_path)
                col_name = f'temp{temp}'
                mapper_df.at[idx, col_name] = summary_content
    
    print(f"\n✅ Found matches: {found_count}")
    print(f"⚠️ Missing matches: {missing_count}")
    
    # Remove temporary column
    mapper_df = mapper_df.drop(columns=['video_id'])
    
    # Save result
    print(f"\n💾 Saving to {OUTPUT_FILE}...")
    mapper_df.to_excel(OUTPUT_FILE, index=False)
    print("✅ Done!")
    
    # Show sample
    print("\n📋 Sample output:")
    print(mapper_df.head(3)[['Category', 'Duration(min)', 'temp0.0', 'temp0.2']].to_string())

if __name__ == "__main__":
    main()
