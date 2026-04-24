#!/usr/bin/env python3
import os
import subprocess
import sys

# Directory containing unit config files
config_dir = 'api_transition/configs/units'

# Get list of YAML files, excluding template and non-config files
config_files = [f for f in os.listdir(config_dir) if f.endswith('.yaml') and not f.startswith('_')]

# Sort for consistent order
config_files.sort()

print(f"Found {len(config_files)} unit configurations to process.")

for config in config_files:
    config_path = f'{config_dir}/{config}'
    command = ['python3', '-u', '-m', 'api_transition.full_pipeline', '--config', config_path]
    
    print(f"\nRunning for {config}...")
    try:
        result = subprocess.run(command, check=True, capture_output=True, text=True)
        print(f"Completed {config} successfully.")
        # Optionally print stdout if needed
        # print(result.stdout)
    except subprocess.CalledProcessError as e:
        print(f"Error running {config}: {e}")
        print(f"Stdout: {e.stdout}")
        print(f"Stderr: {e.stderr}")
        # Continue to next or stop? For now, continue
        continue

print("\nAll units processed.")