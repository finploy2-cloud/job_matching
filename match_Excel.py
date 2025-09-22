import pandas as pd
import os
from getpass import getpass  # Not needed for Excel, but keeping structure

# =============================================================================
# CONFIGURATION - UPDATE THESE!
# =============================================================================

# File paths
JOBS_FILE = r'D:\FINPLOY\Matching codes\matching_betwen_Excel_file\job_id_active.xlsx'      # Your jobs Excel file
CANDIDATES_FILE = r'D:\FINPLOY\Matching codes\matching_betwen_Excel_file\linup.xlsx'         # Your candidates Excel file
OUTPUT_DIR = r'D:\FINPLOY\Matching codes\matching_betwen_Excel_file\matchfiles'

# COLUMN NAMES - UPDATE THESE TO MATCH YOUR EXCEL FILES!
JOBS_COLUMNS = {
    'job_id_col': 'job_id',           # Column name for job ID in jobs.xlsx
    'composite_key_col': 'composit_key'  # Column name for composite key in jobs.xlsx
}

CANDIDATES_COLUMNS = {
    'candidate_id_col': 'candidate_id',   # Column name for candidate ID in candidates.xlsx
    'composite_key_col': 'composit_key'   # Column name for composite key in candidates.xlsx
}

# =============================================================================
def load_excel_data():
    """Load jobs and candidates from Excel files."""
    try:
        # Load jobs file (assumes columns: job_id, composit_key, and other job details)
        jobs_df = pd.read_excel(JOBS_FILE)
        print(f"✅ Loaded {len(jobs_df)} jobs from {JOBS_FILE}")
        
        # Load candidates file (assumes columns: candidate_id, composit_key, and other candidate details)
        candidates_df = pd.read_excel(CANDIDATES_FILE)
        print(f"✅ Loaded {len(candidates_df)} candidates from {CANDIDATES_FILE}")
        
        return jobs_df, candidates_df
        
    except FileNotFoundError as e:
        print(f"❌ File not found: {e}")
        print("   Make sure jobs.xlsx and candidates.xlsx are in the same folder as this script")
        return None, None
    except Exception as e:
        print(f"❌ Error loading Excel files: {e}")
        return None, None

def parse_composite_key(key_str):
    """Split key into parts; return prefix, salary, and full parts."""
    if pd.isna(key_str) or '_' not in str(key_str) or str(key_str).count('_') != 3:
        raise ValueError(f"Invalid key '{key_str}': Must be exactly 4 parts separated by '_' (e.g., '126_5_8_2.6').")
    
    parts = str(key_str).split('_')
    prefix = '_'.join(parts[:3])  # First 3: locationcode_subproduct_product
    try:
        salary = float(parts[3])  # Extract salary from composite_key
    except ValueError:
        raise ValueError(f"Invalid salary '{parts[3]}' in key '{key_str}': Must be a number like 2.6.")
    
    return prefix, salary, parts

def find_matching_candidates_for_all_jobs(jobs_df, candidates_df):
    """Find matching candidates for ALL jobs at once - salary extracted from composite_key."""
    results = {}
    
    # Create a prefix lookup for candidates (for faster matching)
    candidates_by_prefix = {}
    for _, row in candidates_df.iterrows():
        cand_key = row[CANDIDATES_COLUMNS['composite_key_col']]
        try:
            prefix, cand_salary, _ = parse_composite_key(cand_key)
            if prefix not in candidates_by_prefix:
                candidates_by_prefix[prefix] = []
            candidates_by_prefix[prefix].append({
                'candidate_id': row[CANDIDATES_COLUMNS['candidate_id_col']],
                'salary': cand_salary,  # Extracted from composite_key
                'full_row': row
            })
            print(f"📝 Processed candidate {row[CANDIDATES_COLUMNS['candidate_id_col']]}: prefix='{prefix}', salary={cand_salary}")
        except ValueError as e:
            print(f"⚠️ Skipping malformed candidate key '{cand_key}': {e}")
            continue  # Skip malformed keys
    
    print(f"📊 Created prefix lookup with {len(candidates_by_prefix)} unique prefixes")
    
    # Process each job
    for _, job_row in jobs_df.iterrows():
        job_id = job_row[JOBS_COLUMNS['job_id_col']]
        job_key = job_row[JOBS_COLUMNS['composite_key_col']]
        
        try:
            prefix, target_salary, _ = parse_composite_key(job_key)
            print(f"\n🔍 Processing Job {job_id} with key '{job_key}' (salary ≤ {target_salary})")
            
            # Find candidates with matching prefix
            if prefix in candidates_by_prefix:
                matches = []
                candidate_list = candidates_by_prefix[prefix]
                
                for cand_data in candidate_list:
                    if cand_data['salary'] <= target_salary:
                        matches.append(cand_data['full_row'])
                
                if matches:
                    match_df = pd.DataFrame(matches)
                    results[job_id] = {
                        'job_key': job_key,
                        'target_salary': target_salary,  # From composite_key
                        'matches': match_df,
                        'count': len(match_df)
                    }
                    print(f"✅ Found {len(match_df)} matching candidates for Job {job_id}")
                else:
                    print(f"❌ No candidates with salary ≤ {target_salary} for Job {job_id}")
                    results[job_id] = {
                        'job_key': job_key,
                        'target_salary': target_salary,
                        'matches': pd.DataFrame(),
                        'count': 0
                    }
            else:
                print(f"❌ No candidates found for prefix '{prefix}' (Job {job_id})")
                # Still parse the salary for consistency
                _, target_salary, _ = parse_composite_key(job_key)
                results[job_id] = {
                    'job_key': job_key,
                    'target_salary': target_salary,
                    'matches': pd.DataFrame(),
                    'count': 0
                }
                
        except ValueError as e:
            print(f"🚫 Error processing Job {job_id}: {e}")
            continue
    
    return results

def export_to_single_excel(results):
    """Export all matches to a single Excel file with job_id column added."""
    if not os.path.exists(OUTPUT_DIR):
        try:
            os.makedirs(OUTPUT_DIR)
            print(f"\n📁 Created directory: {OUTPUT_DIR}")
        except OSError as e:
            print(f"❌ Failed to create directory {OUTPUT_DIR}: {e}")
            return 0
    
    all_matches_dfs = []
    total_matches = 0
    
    for job_id, data in results.items():
        match_df = data['matches']
        
        if not match_df.empty:
            # Add job_id column to the matches DataFrame
            match_df_with_job = match_df.copy()
            match_df_with_job[JOBS_COLUMNS['job_id_col']] = job_id  # Add job_id column at the beginning or end; adjust position if needed
            
            # Move job_id to first column for better readability
            cols = [JOBS_COLUMNS['job_id_col']] + [col for col in match_df_with_job.columns if col != JOBS_COLUMNS['job_id_col']]
            match_df_with_job = match_df_with_job[cols]
            
            all_matches_dfs.append(match_df_with_job)
            total_matches += len(match_df)
            
            print(f"\n📊 Job {job_id} Matches Preview ({len(match_df)} candidates):")
            # Show key info extracted from composite_key for preview
            match_df_with_job['extracted_salary'] = match_df_with_job[CANDIDATES_COLUMNS['composite_key_col']].apply(
                lambda x: parse_composite_key(x)[1] if pd.notna(x) else None
            )
            preview_cols = [JOBS_COLUMNS['job_id_col'], CANDIDATES_COLUMNS['candidate_id_col'], CANDIDATES_COLUMNS['composite_key_col'], 'extracted_salary']
            print(match_df_with_job[preview_cols].head().to_string(index=False))
        else:
            print(f"📝 Job {job_id}: No matches (skipping)")
    
    if all_matches_dfs:
        # Concatenate all matching DataFrames
        combined_df = pd.concat(all_matches_dfs, ignore_index=True)
        
        # Define output file
        output_file = os.path.join(OUTPUT_DIR, 'all_job_candidate_matches.xlsx')
        
        try:
            # Export to single Excel file
            combined_df.to_excel(output_file, index=False, engine='openpyxl')
            print(f"\n💾 Exported {total_matches} total matches to single Excel: {output_file}")
            
            # Verify file creation
            if os.path.exists(output_file):
                print(f"✅ File verified: {output_file}")
                
        except Exception as e:
            print(f"❌ Failed to export combined Excel: {e}")
            return 0
    else:
        print("\n⚠️ No matches found across all jobs - no file created.")
        return 0
    
    return total_matches

def main():
    """Main function - loads Excel files and processes all jobs."""
    print("🚀 Excel-based Job-Candidate Matcher (Salary from Composite Key)")
    print("=" * 50)
    
    # Load Excel data
    jobs_df, candidates_df = load_excel_data()
    if jobs_df is None or candidates_df is None:
        return
    
    # Find matches for all jobs
    print("\n🔍 Finding matches for all jobs...")
    results = find_matching_candidates_for_all_jobs(jobs_df, candidates_df)
    
    if not results:
        print("❌ No results to process")
        return
    
    # Export results to single Excel
    print("\n💾 Exporting results to single Excel file...")
    total_exported = export_to_single_excel(results)
    
    # Summary report
    print("\n📈 SUMMARY REPORT:")
    print("-" * 30)
    for job_id, data in results.items():
        status = "✅ HAS MATCHES" if data['count'] > 0 else "❌ NO MATCHES"
        print(f"Job {job_id}: {status} ({data['count']} candidates, target ≤ {data['target_salary']})")
    
    print(f"\n🎯 All done! Check {OUTPUT_DIR} for 'all_job_candidate_matches.xlsx'.")
    print(f"📁 Total rows in Excel: {total_exported}")

if __name__ == '__main__':
    main()