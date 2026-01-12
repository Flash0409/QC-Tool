#!/usr/bin/env python3
"""
Migration Script: Convert old categories.json to new format with reference numbers

Usage:
    python migrate_categories.py [old_categories.json] [output_path]
"""

import json
import os
import sys
from datetime import datetime


def generate_ref_number(counter):
    """Generate numeric reference number like 01, 02, ..., 99"""
    return f"{counter:02d}"


def migrate_categories(old_path, output_path):
    """Migrate old categories to new format with numeric auto reference numbers (01-99)"""
    
    # Load old categories
    try:
        with open(old_path, 'r', encoding='utf-8') as f:
            old_categories = json.load(f)
    except FileNotFoundError:
        print(f"❌ Old categories file not found: {old_path}")
        print("📝 Creating new categories from template...")
        old_categories = []
    except json.JSONDecodeError as e:
        print(f"❌ Error parsing old categories: {e}")
        return False
    
    # New structure - NUMERIC ONLY
    new_structure = {
        "version": "2.0",
        "description": "Categories with automatic reference number assignment (numeric 01-99)",
        "migrated_from": old_path if os.path.exists(old_path) else None,
        "migration_date": datetime.now().isoformat(),
        "categories": [],
        "custom_category": {
            "name": "Custom",
            "ref_counter": 50,  # Custom starts at 50
            "allow_custom_ref": True,
            "note": "Custom punches start at 50 and increment"
        }
    }
    
    # Global counter for reference numbers (01-49 for categories, 50-99 for custom)
    global_counter = 1
    
    # Migrate each category
    for old_cat in old_categories:
        cat_name = old_cat.get("name", "Unknown")
        mode = old_cat.get("mode", "parent")
        
        new_cat = {
            "name": cat_name,
            "mode": mode,
            "ref_counter": 1  # Not used for fixed categories, kept for compatibility
        }
        
        # Handle template mode
        if mode == "template":
            new_cat["ref_number"] = generate_ref_number(global_counter)
            new_cat["template"] = old_cat.get("template", "")
            new_cat["inputs"] = old_cat.get("inputs", [])
            global_counter += 1
        
        # Handle parent mode with subcategories
        elif mode == "parent" and "subcategories" in old_cat:
            new_cat["subcategories"] = []
            
            for old_sub in old_cat["subcategories"]:
                if global_counter >= 50:
                    print(f"⚠️ Warning: Exceeded 49 category items. Remaining will use custom range.")
                    break
                    
                new_sub = {
                    "name": old_sub.get("name", "Unknown"),
                    "ref_number": generate_ref_number(global_counter),
                    "template": old_sub.get("template", ""),
                    "inputs": old_sub.get("inputs", [])
                }
                new_cat["subcategories"].append(new_sub)
                global_counter += 1
        
        new_structure["categories"].append(new_cat)
    
    # Save migrated structure
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(new_structure, f, indent=2, ensure_ascii=False)
        
        print(f"✅ Migration successful!")
        print(f"📁 Output: {output_path}")
        print(f"📊 Migrated {len(new_structure['categories'])} categories")
        
        # Print summary
        print("\n📋 Reference Number Assignments:")
        for cat in new_structure["categories"]:
            if cat["mode"] == "parent" and "subcategories" in cat:
                refs = [sub["ref_number"] for sub in cat["subcategories"]]
                if refs:
                    print(f"  {cat['name']:30} → {refs[0]}-{refs[-1]} ({len(refs)} items)")
            elif "ref_number" in cat:
                print(f"  {cat['name']:30} → {cat['ref_number']}")
        
        print(f"\n  Custom Category               → 50-99 (auto-increment)")
        print(f"\n⚠️ Note: Reference numbers 01-49 are for categories")
        print(f"         Reference numbers 50-99 are for custom punches")
        
        return True
        
    except Exception as e:
        print(f"❌ Error saving migrated categories: {e}")
        return False


def create_backup(filepath):
    """Create backup of existing file"""
    if not os.path.exists(filepath):
        return None
    
    backup_path = f"{filepath}.backup.{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    try:
        with open(filepath, 'r') as src:
            with open(backup_path, 'w') as dst:
                dst.write(src.read())
        print(f"📦 Backup created: {backup_path}")
        return backup_path
    except Exception as e:
        print(f"⚠️ Could not create backup: {e}")
        return None


def main():
    """Main migration entry point"""
    
    # Default paths
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_old_path = os.path.join(script_dir, "..", "assets", "categories.json")
    default_output_path = os.path.join(script_dir, "..", "assets", "categories.json")
    
    # Parse arguments
    if len(sys.argv) > 1:
        old_path = sys.argv[1]
    else:
        old_path = default_old_path
    
    if len(sys.argv) > 2:
        output_path = sys.argv[2]
    else:
        output_path = default_output_path
    
    print("=" * 60)
    print("  CATEGORY MIGRATION TOOL v2.0")
    print("  Auto Reference Number Assignment")
    print("=" * 60)
    print(f"\n📂 Input:  {old_path}")
    print(f"📂 Output: {output_path}")
    
    # Confirm if overwriting
    if os.path.exists(output_path) and output_path == old_path:
        response = input("\n⚠️  This will OVERWRITE the existing categories file.\n   Continue? (yes/no): ")
        if response.lower() not in ('yes', 'y'):
            print("❌ Migration cancelled.")
            return
        
        # Create backup
        create_backup(output_path)
    
    print("\n🔄 Starting migration...\n")
    
    # Run migration
    success = migrate_categories(old_path, output_path)
    
    if success:
        print("\n✅ Migration complete!")
        print("\n📝 Next steps:")
        print("   1. Review the migrated categories.json")
        print("   2. Restart the Quality Inspection Tool")
        print("   3. Reference numbers will now auto-populate")
    else:
        print("\n❌ Migration failed. Check errors above.")
    
    print("=" * 60)


if __name__ == "__main__":
    main()
