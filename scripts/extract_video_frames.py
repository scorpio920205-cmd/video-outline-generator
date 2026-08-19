#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Fast Video Frame Extractor Utility using FFmpeg fast-seeking.
Extracts frames at regular intervals and saves them with timestamps in filename.
"""

import os
import sys
import time
import argparse
import subprocess
import json

def get_video_duration(video_path):
    cmd = [
        'ffprobe', '-v', 'quiet', '-print_format', 'json',
        '-show_format', video_path
    ]
    res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
    if res.returncode == 0:
        data = json.loads(res.stdout)
        return float(data.get('format', {}).get('duration', 0))
    return 0.0

def format_ts(seconds):
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    return f"{h:02d}_{m:02d}_{s:02d}"

def extract_frames(video_path, output_dir, interval_seconds=60, scale_width=1280):
    os.makedirs(output_dir, exist_ok=True)
    duration = get_video_duration(video_path)
    if duration <= 0:
        print(f"Error: Unable to detect duration for {video_path}")
        return False
    
    timestamps = list(range(0, int(duration), interval_seconds))
    total_frames = len(timestamps)
    print(f"Video Duration: {duration:.1f}s ({duration/60:.1f} mins)")
    print(f"Extracting {total_frames} frames (1 frame every {interval_seconds}s)...")
    
    t0 = time.time()
    extracted = 0
    for idx, sec in enumerate(timestamps):
        ts_str = format_ts(sec)
        out_file = os.path.join(output_dir, f"frame_{idx+1:03d}_{ts_str}.jpg")
        
        if not os.path.exists(out_file):
            cmd = [
                'ffmpeg', '-ss', str(sec), '-i', video_path,
                '-vframes', '1', '-q:v', '3',
                '-vf', f'scale={scale_width}:-1',
                out_file, '-y'
            ]
            subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        extracted += 1
        
        if (idx + 1) % 20 == 0 or (idx + 1) == total_frames:
            elapsed = time.time() - t0
            print(f"Progress: {idx+1}/{total_frames} frames ({elapsed:.1f}s)", flush=True)
            
    print(f"Completed! Extracted {extracted} frames in {time.time()-t0:.1f}s to '{output_dir}'.")
    return True

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Extract video frames at fixed intervals using fast seeking.")
    parser.add_argument('--video', required=True, help="Path to input video file")
    parser.add_argument('--output', default="./frames", help="Output directory for frames")
    parser.add_argument('--interval', type=int, default=60, help="Interval in seconds between frames (default: 60)")
    parser.add_argument('--scale', type=int, default=1280, help="Output image width in pixels (default: 1280)")
    
    args = parser.parse_args()
    extract_frames(args.video, args.output, args.interval, args.scale)
