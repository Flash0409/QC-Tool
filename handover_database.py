"""
Handover Database Manager
Manages Quality <-> Production handover workflow
"""

import json
import os
from datetime import datetime
from typing import List, Dict, Optional


class HandoverDB:
    """Manages handover records between Quality and Production"""
    
    def __init__(self, db_path: str = None):
        """Initialize database at specified path"""
        if db_path is None:
            db_path = os.path.join(os.path.dirname(__file__), "handover_db.json")
        
        self.db_path = db_path
        self._ensure_db_exists()
    
    def _ensure_db_exists(self):
        """Create database file if it doesn't exist"""
        if not os.path.exists(self.db_path):
            self._save_db({
                "quality_to_production": [],
                "production_to_quality": []
            })
    
    def _load_db(self) -> dict:
        """Load database from file"""
        try:
            with open(self.db_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading database: {e}")
            return {
                "quality_to_production": [],
                "production_to_quality": []
            }
    
    def _save_db(self, data: dict):
        """Save database to file"""
        with open(self.db_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    
    # ================================================================
    # QUALITY TO PRODUCTION HANDOVER
    # ================================================================
    
    def add_quality_handover(self, handover_data: dict) -> bool:
        """
        Add a new Quality -> Production handover
        
        Args:
            handover_data: Dict containing:
                - cabinet_id: str
                - project_name: str
                - sales_order_no: str
                - pdf_path: str
                - excel_path: str
                - session_path: str
                - total_punches: int
                - open_punches: int
                - closed_punches: int
                - handed_over_by: str
                - handed_over_date: str (ISO format)
        """
        try:
            db = self._load_db()
            
            # Check if already handed over
            existing = next(
                (item for item in db["quality_to_production"] 
                 if item["cabinet_id"] == handover_data["cabinet_id"] 
                 and item["status"] == "pending"),
                None
            )
            
            if existing:
                return False  # Already in production queue
            
            # Add status and timestamp
            handover_data["status"] = "pending"  # pending, in_progress, completed
            handover_data["received_by"] = None
            handover_data["received_date"] = None
            handover_data["completed_by"] = None
            handover_data["completed_date"] = None
            
            db["quality_to_production"].append(handover_data)
            self._save_db(db)
            return True
            
        except Exception as e:
            print(f"Error adding quality handover: {e}")
            return False
    
    def get_pending_production_items(self) -> List[Dict]:
        """Get all items pending in production"""
        db = self._load_db()
        return [
            item for item in db["quality_to_production"]
            if item["status"] in ["pending", "in_progress"]
        ]
    
    def update_production_status(self, cabinet_id: str, status: str, user: str = None):
        """Update production status for an item"""
        db = self._load_db()
        
        for item in db["quality_to_production"]:
            if item["cabinet_id"] == cabinet_id:
                item["status"] = status
                
                if status == "in_progress" and not item.get("received_by"):
                    item["received_by"] = user
                    item["received_date"] = datetime.now().isoformat()
                
                elif status == "completed":
                    item["completed_by"] = user
                    item["completed_date"] = datetime.now().isoformat()
                
                self._save_db(db)
                return True
        
        return False
    
    # ================================================================
    # PRODUCTION TO QUALITY HANDBACK
    # ================================================================
    
    def add_production_handback(self, handback_data: dict) -> bool:
        """
        Add a Production -> Quality handback
        
        Args:
            handback_data: Dict containing:
                - cabinet_id: str
                - project_name: str
                - sales_order_no: str
                - pdf_path: str
                - excel_path: str
                - session_path: str
                - rework_completed_by: str
                - rework_completed_date: str
                - production_remarks: str (optional)
        """
        try:
            db = self._load_db()
            
            # Mark quality handover as completed
            for item in db["quality_to_production"]:
                if item["cabinet_id"] == handback_data["cabinet_id"]:
                    item["status"] = "completed"
                    break
            
            # Add handback status
            handback_data["status"] = "pending"  # pending, verified, closed
            handback_data["verified_by"] = None
            handback_data["verified_date"] = None
            
            db["production_to_quality"].append(handback_data)
            self._save_db(db)
            return True
            
        except Exception as e:
            print(f"Error adding production handback: {e}")
            return False
    
    def get_pending_quality_items(self) -> List[Dict]:
        """Get all items pending quality verification"""
        db = self._load_db()
        return [
            item for item in db["production_to_quality"]
            if item["status"] == "pending"
        ]
    
    def update_quality_verification(self, cabinet_id: str, status: str, user: str = None):
        """Update quality verification status"""
        db = self._load_db()
        
        for item in db["production_to_quality"]:
            if item["cabinet_id"] == cabinet_id:
                item["status"] = status
                item["verified_by"] = user
                item["verified_date"] = datetime.now().isoformat()
                
                self._save_db(db)
                return True
        
        return False
    
    # ================================================================
    # UTILITY FUNCTIONS
    # ================================================================
    
    def get_item_by_cabinet_id(self, cabinet_id: str, queue: str = "quality_to_production") -> Optional[Dict]:
        """Get handover item by cabinet ID"""
        db = self._load_db()
        
        for item in db.get(queue, []):
            if item["cabinet_id"] == cabinet_id:
                return item
        
        return None
    
    def get_all_handovers(self) -> Dict:
        """Get complete database"""
        return self._load_db()
    
    def cleanup_completed(self, days_old: int = 30):
        """Remove completed items older than specified days"""
        db = self._load_db()
        cutoff = datetime.now().timestamp() - (days_old * 24 * 60 * 60)
        
        # Clean quality_to_production
        db["quality_to_production"] = [
            item for item in db["quality_to_production"]
            if item["status"] != "completed" or
            datetime.fromisoformat(item["completed_date"]).timestamp() > cutoff
        ]
        
        # Clean production_to_quality
        db["production_to_quality"] = [
            item for item in db["production_to_quality"]
            if item["status"] != "closed" or
            datetime.fromisoformat(item["verified_date"]).timestamp() > cutoff
        ]
        
        self._save_db(db)
